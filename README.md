# Capstone Project Assignment Tool

A web app that automatically assigns students to capstone projects based on their ranked preferences, while respecting seat caps and per-major quotas.

## How it works

1. Upload a **Student File** (`.xlsx`) with columns:
   - `Student Name`, `Major`, `Rank 1`, `Rank 2`, `Rank 3`, `Rank 4`, `Rank 5`
   - Each rank column contains the project name the student wants for that preference.

2. Upload a **Projects File** (`.xlsx`) with columns:
   - `Project Name`, `Total Seats`, and one column per major (e.g. `CS Seats`, `EE Seats`)

3. The tool runs a greedy assignment algorithm: iterates from Rank 1 → 5 across all students, placing each student in their highest available preference that still has room (both overall and for their major).

4. Download the results Excel file — one sheet per project listing assigned students (name, major, preference rank received), plus an "Unassigned Students" sheet.

## Running locally

```bash
pip install -r requirements.txt
python app.py
```

Then open [http://localhost:3000](http://localhost:3000).

## Project structure

```
app.py                  # Flask backend + assignment algorithm
templates/index.html    # Frontend UI (Jinja2 template)
requirements.txt        # Python dependencies
```
