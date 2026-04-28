import os
import io
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename

# Initialize Flask app
app = Flask(__name__)

# Limit upload size to 16 MB to prevent abuse / accidental large uploads
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

# Only allow Excel files
ALLOWED_EXTENSIONS = {'xlsx'}


def allowed_file(filename):
    """Check file extension against allowed types."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def parse_students_excel(file_bytes):
    """
    Reads student Excel file and converts it into a structured list of student dictionaries.
    Expected columns:
        - Student Name
        - Major
        - Rank/Preference columns (various supported naming styles)
    """

    df = pd.read_excel(io.BytesIO(file_bytes))

    # Normalize column names (remove whitespace issues)
    df.columns = [str(c).strip() for c in df.columns]

    # Required fields validation
    required = ['Student Name', 'Major']
    for col in required:
        if col not in df.columns:
            raise ValueError(f"Student file is missing the '{col}' column.")

    # Try to detect preference/rank columns flexibly (Rank 1, Choice 1, etc.)
    rank_cols = []
    for i in range(1, 6):
        candidates = [
            c for c in df.columns
            if c.lower().replace(' ', '') in [
                f'rank{i}', f'choice{i}', f'preference{i}',
                f'rank{i}project', f'project{i}'
            ]
        ]
        if candidates:
            rank_cols.append(candidates[0])
        else:
            # fallback heuristic search
            for c in df.columns:
                if str(i) in c and 'rank' in c.lower():
                    rank_cols.append(c)
                    break

    # Final fallback format: "Rank 1", "Rank 2", etc.
    if not rank_cols:
        for i in range(1, 6):
            col_name = f'Rank {i}'
            if col_name in df.columns:
                rank_cols.append(col_name)

    if len(rank_cols) < 1:
        raise ValueError(
            "Student file must have at least one preference column."
        )

    students = []
    row_index = 0

    # Convert dataframe rows into structured student objects
    for _, row in df.iterrows():
        name = str(row['Student Name']).strip() if pd.notna(row['Student Name']) else ''
        major = str(row['Major']).strip() if pd.notna(row['Major']) else ''

        # Skip invalid rows
        if not name or name.lower() == 'nan':
            continue

        # Collect ranked preferences in order
        prefs = []
        for col in rank_cols:
            val = row.get(col, None)
            if pd.notna(val) and str(val).strip().lower() != 'nan':
                prefs.append(str(val).strip())

        students.append({
            'id': row_index,
            'name': name,
            'major': major,
            'preferences': prefs
        })
        row_index += 1

    return students


def parse_projects_excel(file_bytes):
    """
    Reads project Excel file and converts it into structured project dictionaries.
    Supports:
        - Project name column (flexible naming)
        - Total seat capacity
        - Optional major-specific quotas
    """

    df = pd.read_excel(io.BytesIO(file_bytes))
    df.columns = [str(c).strip() for c in df.columns]

    # Detect project name column
    name_col = None
    for c in df.columns:
        if c.lower() in ['project name', 'capstone name', 'project', 'capstone', 'name']:
            name_col = c
            break
    if name_col is None:
        raise ValueError("Projects file must have a 'Project Name' column.")

    # Detect seat/capacity column
    seats_col = None
    for c in df.columns:
        if c.lower() in ['total seats', 'seats', 'total', 'capacity', 'group size', 'number of people']:
            seats_col = c
            break
    if seats_col is None:
        raise ValueError("Projects file must have a 'Total Seats' column.")

    skip_cols = {name_col.lower(), seats_col.lower()}
    major_cols = [c for c in df.columns if c.lower() not in skip_cols]

    projects = []

    for _, row in df.iterrows():
        pname = str(row[name_col]).strip() if pd.notna(row[name_col]) else ''
        if not pname or pname.lower() == 'nan':
            continue

        # Parse total capacity safely
        try:
            total_seats = int(row[seats_col]) if pd.notna(row[seats_col]) else 0
        except (ValueError, TypeError):
            total_seats = 0

        # Extract optional major quotas per project
        major_quotas = {}
        for mc in major_cols:
            val = row.get(mc, None)
            if pd.notna(val):
                try:
                    quota = int(val)
                    if quota > 0:
                        major_name = mc.replace(' Seats', '').replace(' seats', '').strip()
                        major_quotas[major_name] = quota
                except (ValueError, TypeError):
                    pass

        projects.append({
            'name': pname,
            'total_seats': total_seats,
            'major_quotas': major_quotas,
            'assigned': []
        })

    return projects


def assign_students(students, projects):
    """
    Core matching algorithm:
    - Iterates by preference rank (fairness)
    - Assigns students if capacity + major constraints allow

    This algorithm is a greedy, constraint-based assignment method rather than 
    a mixed-integer programming model or genetic algorithm. It assigns students 
    to projects in two phases: first, it sweeps through preference ranks (Rank 1 
    to Rank 5), repeatedly trying to place unassigned students into their highest 
    available choice while respecting project capacity and optional major-specific 
    quotas; once a student is assigned, they are immediately locked in. In the 
    second pass, any remaining unassigned students are given another chance to be 
    placed into any project in their preference list. The approach is fast and 
    deterministic, but it does not guarantee a globally optimal solution because 
    assignments are made locally at each step rather than through a global 
    optimization process.
    """
    
    project_map = {p['name']: p for p in projects}

    # Track remaining capacity
    project_seats_remaining = {p['name']: p['total_seats'] for p in projects}

    # Track remaining major-specific slots
    project_major_remaining = {
        p['name']: dict(p['major_quotas']) for p in projects
    }

    def can_assign(student, project_name):
        """Check if student can be assigned to a project."""
        if project_name not in project_map:
            return False
        if project_seats_remaining[project_name] <= 0:
            return False

        major = student['major']
        major_rem = project_major_remaining[project_name]

        # No major restrictions means open slot
        if not major_rem:
            return True

        # Case-insensitive major matching
        normalized = {k.lower(): k for k in major_rem}
        key = normalized.get(major.lower())

        if key is None:
            return False

        return major_rem[key] > 0

    def do_assign(student, project_name, rank):
        """Perform assignment and update capacity tracking."""
        project_seats_remaining[project_name] -= 1

        major = student['major']
        major_rem = project_major_remaining[project_name]

        if major_rem:
            normalized = {k.lower(): k for k in major_rem}
            key = normalized.get(major.lower())
            if key:
                project_major_remaining[project_name][key] -= 1

        project_map[project_name]['assigned'].append({
            'name': student['name'],
            'major': student['major'],
            'rank': rank
        })

    assigned_ids = set()
    unassigned = []

    # Pass 1: strict rank-by-rank assignment (fair allocation)
    for rank_idx in range(5):
        for student in students:
            if student['id'] in assigned_ids:
                continue
            if rank_idx >= len(student['preferences']):
                continue

            pref = student['preferences'][rank_idx]
            if can_assign(student, pref):
                do_assign(student, pref, rank_idx + 1)
                assigned_ids.add(student['id'])

    # Pass 2: fallback assignment (try any remaining preference)
    for student in students:
        if student['id'] not in assigned_ids:
            for rank_idx, pref in enumerate(student['preferences']):
                if can_assign(student, pref):
                    do_assign(student, pref, rank_idx + 1)
                    assigned_ids.add(student['id'])
                    break
            if student['id'] not in assigned_ids:
                unassigned.append(student)

    return projects, unassigned


def generate_output_excel(projects, unassigned):
    """
    Builds formatted Excel output:
        - Summary sheet
        - One sheet per project
        - Unassigned students sheet
    """

    wb = openpyxl.Workbook()

    # Styling presets
    header_font = Font(bold=True, color='FFFFFF', size=11)
    header_fill = PatternFill(start_color='2563EB', end_color='2563EB', fill_type='solid')
    alt_fill = PatternFill(start_color='EFF6FF', end_color='EFF6FF', fill_type='solid')
    center = Alignment(horizontal='center', vertical='center')
    left = Alignment(horizontal='left', vertical='center')

    thin = Side(style='thin', color='CBD5E1')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ---------------- SUMMARY SHEET ----------------
    summary_ws = wb.active
    summary_ws.title = 'Summary'

    summary_ws.column_dimensions['A'].width = 30

    headers = ['Project Name', 'Total Assigned', 'Total Seats', 'Seats Filled %']

    for col, val in enumerate(headers, 1):
        cell = summary_ws.cell(row=1, column=col, value=val)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    for row_i, p in enumerate(projects, 2):
        assigned_count = len(p['assigned'])
        total = p['total_seats']
        pct = f"{int(assigned_count / total * 100)}%" if total > 0 else 'N/A'

        for col, val in enumerate([p['name'], assigned_count, total, pct], 1):
            cell = summary_ws.cell(row=row_i, column=col, value=val)
            cell.alignment = center if col > 1 else left
            cell.border = border
            if row_i % 2 == 0:
                cell.fill = alt_fill

    # ---------------- PROJECT SHEETS ----------------
    for p in projects:
        ws = wb.create_sheet(title=p['name'][:31])

        ws.cell(row=1, column=1, value=f"Project: {p['name']}")
        ws.merge_cells('A1:C1')

        # headers
        for col, h in enumerate(['Student Name', 'Major', 'Preference Rank'], 1):
            cell = ws.cell(row=2, column=col, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center

        # assigned students
        for row_i, student in enumerate(p['assigned'], 3):
            ws.append([student['name'], student['major'], f"#{student['rank']} choice"])

    # ---------------- UNASSIGNED SHEET ----------------
    if unassigned:
        ws = wb.create_sheet(title='Unassigned Students')

        ws.append(['Student Name', 'Major', 'Preferences Listed'])

        for s in unassigned:
            ws.append([s['name'], s['major'], ', '.join(s['preferences'])])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def generate_student_template():
    """Creates downloadable student Excel template with example rows."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Students'

    headers = ['Student Name', 'Major', 'Rank 1', 'Rank 2', 'Rank 3', 'Rank 4', 'Rank 5']
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def generate_projects_template():
    """Creates downloadable project Excel template with example structure."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Projects'

    headers = ['Project Name', 'Total Seats']
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ---------------- ROUTES ----------------

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/assign', methods=['POST'])
def assign():
    """
    Main API endpoint:
    - Accepts student + project Excel files
    - Runs matching algorithm
    - Returns JSON + base64 Excel output
    """

    if 'students_file' not in request.files or 'projects_file' not in request.files:
        return jsonify({'error': 'Both files are required.'}), 400

    students_file = request.files['students_file']
    projects_file = request.files['projects_file']

    if not students_file.filename or not projects_file.filename:
        return jsonify({'error': 'Both files must be selected.'}), 400

    if not allowed_file(students_file.filename):
        return jsonify({'error': 'Students file must be an Excel (.xlsx) file.'}), 400
    if not allowed_file(projects_file.filename):
        return jsonify({'error': 'Projects file must be an Excel (.xlsx) file.'}), 400

    try:
        students = parse_students_excel(students_file.read())
        projects = parse_projects_excel(projects_file.read())

        if not students:
            return jsonify({'error': 'No valid students found.'}), 400
        if not projects:
            return jsonify({'error': 'No valid projects found.'}), 400

        projects, unassigned = assign_students(students, projects)
        output = generate_output_excel(projects, unassigned)

        # Build API summary response
        summary = {
            'total_students': len(students),
            'assigned': sum(len(p['assigned']) for p in projects),
            'unassigned': len(unassigned),
            'projects': [
                {
                    'name': p['name'],
                    'assigned': len(p['assigned']),
                    'total_seats': p['total_seats']
                }
                for p in projects
            ]
        }

        import base64
        excel_b64 = base64.b64encode(output.read()).decode('utf-8')

        return jsonify({'summary': summary, 'file': excel_b64})

    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    except Exception as e:
        return jsonify({'error': f'Unexpected error: {str(e)}'}), 500


# ---------------- TEMPLATE DOWNLOAD ROUTES ----------------

@app.route('/template/students')
def download_students_template():
    return send_file(
        generate_student_template(),
        as_attachment=True,
        download_name='students_template.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/template/projects')
def download_projects_template():
    return send_file(
        generate_projects_template(),
        as_attachment=True,
        download_name='projects_template.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 3000))
    app.run(host='0.0.0.0', port=port, debug=False)