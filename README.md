# 🎫 Ticket Tracker

A lightweight project-management tool written in Python. It keeps all project
tickets in a simple CSV file and generates a formatted **Excel tracker sheet**
with a live summary dashboard — the kind of status report a project manager
shares with the team every week.

## Features

- **Log tickets** from the command line (bugs, features, tasks, improvements)
- **List and filter** tickets by status or assignee, with overdue flags
- **Update** status, priority, owner or due date as work progresses
- **Generate `tracker.xlsx`** with:
  - **Summary sheet** — total / open / done / overdue counts, completion %,
    breakdown by status and priority, assignee workload table, pie and bar charts
  - **Tickets sheet** — every ticket with colour-coded status and priority,
    dropdowns, filters, frozen header and an auto-calculated *Overdue?* column
  - All summary numbers are Excel formulas, so they update when you edit the sheet

## Project structure

```
ticket-tracker/
├── tracker.py        # CLI: add / list / update / report
├── report.py         # builds the Excel tracker sheet
├── tickets.csv       # the ticket "database" (sample data included)
├── tracker.xlsx      # sample generated tracker
├── requirements.txt
└── README.md
```

## Getting started

```bash
git clone https://github.com/<your-username>/ticket-tracker.git
cd ticket-tracker
pip install -r requirements.txt
```

## Usage

```bash
# Create a ticket
python tracker.py add --title "Search returns duplicates" --type Bug \
    --priority High --assignee Riya --due 2026-10-12 \
    --description "Happens when filter + keyword are combined"

# List all tickets, or filter
python tracker.py list
python tracker.py list --status "In Progress"
python tracker.py list --assignee Arjun

# Move a ticket along
python tracker.py update TKT-003 --status "In Review"
python tracker.py update TKT-006 --assignee Karan --priority Medium

# Build the tracker sheet
python tracker.py report                 # -> tracker.xlsx
python tracker.py report --out sprint-5.xlsx
```

### Ticket fields

| Field       | Values                                              |
|-------------|-----------------------------------------------------|
| Type        | Bug, Feature, Task, Improvement                     |
| Priority    | Critical, High, Medium, Low                         |
| Status      | To Do, In Progress, In Review, Blocked, Done        |
| Due date    | `YYYY-MM-DD`                                        |

IDs (`TKT-001`, `TKT-002`, …) and the created date are assigned automatically.

## Workflow

```
To Do ──► In Progress ──► In Review ──► Done
                │
                └──► Blocked (waiting on something)
```

## Ideas for future improvements

- Sprint / milestone field and burndown chart
- Web UI (Flask) on top of the same CSV
- Sync tickets with GitHub Issues

## License

MIT
