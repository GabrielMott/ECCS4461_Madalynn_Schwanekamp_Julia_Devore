# ROBOTS.md

Guidance for AI assistants working on this codebase.

## Project Overview

**Name:** Capstone Project Assignment Tool
**Purpose:** A web app that assigns students to capstone projects based on ranked preferences, respecting per-project seat limits and per-major quotas.
**Stack:** Python 3.11 · Flask · openpyxl · pandas · xlsxwriter · vanilla HTML/CSS/JS frontend (Jinja2 template)

## How It Works

1. User uploads two Excel files:
   - **Students file** — `Student Name`, `Major`, `Rank 1` … `Rank 5`
   - **Projects file** — `Project Name`, `Total Seats`, plus one column per major (e.g. `CS Seats`)
2. Backend parses the files and runs a greedy assignment algorithm:
   - Iterates rank 1 → 5 across all students.
   - Assigns each student to their highest available preference that still has both an open total seat and an open quota for their major.
   - Falls back to next-best available preference for any leftover students.
3. Backend generates a multi-sheet Excel file:
   - **Summary** sheet — project counts, total seats, fill %.
   - One sheet per project — Student Name, Major, Preference Rank.
   - **Unassigned Students** sheet — students who could not be placed.
4. The result is sent back as base64-encoded Excel and offered as a download in-page.

## Project Structure

```
.
├── ROBOTS.md                       # This file — AI guidance
├── replit.md                       # Replit-specific project notes
├── capstone-app.zip                # Packaged source for GitHub export
├── test_students.xlsx              # Sample student input (15 students, 4 majors)
├── test_projects.xlsx              # Sample projects input (5 projects)
│
├── artifacts/
│   ├── capstone-app/               # ← MAIN APP lives here
│   │   ├── app.py                  # Flask backend + assignment algorithm
│   │   ├── templates/
│   │   │   └── index.html          # Single-page UI (upload, results, download)
│   │   └── .replit-artifact/
│   │       └── artifact.toml       # Artifact registration + run command
│   │
│   ├── api-server/                 # Scaffold artifact (not used by this app)
│   └── mockup-sandbox/             # Scaffold artifact (not used by this app)
│
└── .local/                         # Replit internal — do not modify directly
```

## Key Files for AI Edits

| File | What lives here |
|------|-----------------|
| `artifacts/capstone-app/app.py` | All backend logic — Excel parsing, assignment algorithm, output generation, Flask routes, sample-template endpoints |
| `artifacts/capstone-app/templates/index.html` | All frontend — markup, styles, and JS for upload/submit/download flow |
| `artifacts/capstone-app/.replit-artifact/artifact.toml` | Artifact config — port (21101), run command, preview path |

## Important Conventions

- **Python only.** Despite living in a pnpm/TypeScript monorepo, this artifact is pure Python + Jinja2. Do not introduce React, Vite, or Node tooling here.
- **File format: `.xlsx` only.** `xlrd` is not installed, so legacy `.xls` is rejected.
- **Student deduplication uses row ID, not name.** Two students with identical names are treated as distinct people. Do not regress to name-based deduping.
- **Major matching is case-insensitive.** When matching a student's major to a project's major-quota column, lowercase comparison is used.
- **Port comes from the `PORT` env var** (set to 21101 in `artifact.toml`). Do not hardcode.
- **Run command:** `cd artifacts/capstone-app && PORT=21101 python app.py` — managed by the workflow `artifacts/capstone-app: web`.

## Assignment Algorithm Reference

Located in `app.py` → `assign_students(students, projects)`.

```
For rank in 1..5:
    For each unassigned student (in input order):
        candidate = student.preferences[rank-1]
        if project has open total seat AND open quota for student's major:
            assign student → project at this rank
            mark student as assigned

For any student still unassigned:
    Try each of their preferences in order; assign to first available match.
    Otherwise, add to "unassigned" list.
```

The algorithm is intentionally simple and deterministic — order of students in the input file affects ties.

## What NOT to Do

- Do not rewrite the app in another framework or language unless explicitly asked.
- Do not introduce a database — the app is fully stateless and works on in-memory uploads.
- Do not store uploaded files to disk.
- Do not add authentication unless requested — the tool is intended for single-session local use.
- Do not modify `artifacts/api-server/` or `artifacts/mockup-sandbox/` — they are unused scaffold artifacts kept by the monorepo system.
