#!/usr/bin/env python3
"""
Ticket Tracker - a simple project-management tool.

Tickets are stored in tickets.csv. Use this script to add, list and update
tickets, then generate a formatted Excel tracker sheet (tracker.xlsx).

Usage:
    python tracker.py add --title "Login page bug" --type Bug --priority High --assignee Riya --due 2026-10-05
    python tracker.py list [--status "In Progress"] [--assignee Riya]
    python tracker.py update TKT-003 --status Done
    python tracker.py report [--out tracker.xlsx]
"""

import argparse
import csv
import sys
from datetime import date, datetime
from pathlib import Path

TICKETS_FILE = Path(__file__).parent / "tickets.csv"

FIELDS = ["id", "title", "type", "priority", "status", "assignee",
          "created", "due", "description"]
TYPES = ["Bug", "Feature", "Task", "Improvement"]
PRIORITIES = ["Critical", "High", "Medium", "Low"]
STATUSES = ["To Do", "In Progress", "In Review", "Blocked", "Done"]


# ---------------------------------------------------------------- storage
def load_tickets():
    if not TICKETS_FILE.exists():
        return []
    with open(TICKETS_FILE, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def save_tickets(tickets):
    with open(TICKETS_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows(tickets)


def next_id(tickets):
    nums = [int(t["id"].split("-")[1]) for t in tickets if t["id"].startswith("TKT-")]
    return f"TKT-{(max(nums) + 1) if nums else 1:03d}"


def valid_date(value):
    try:
        datetime.strptime(value, "%Y-%m-%d")
        return value
    except ValueError:
        raise argparse.ArgumentTypeError("dates must be YYYY-MM-DD")


def is_overdue(t):
    return (t["status"] != "Done" and t["due"]
            and datetime.strptime(t["due"], "%Y-%m-%d").date() < date.today())


# --------------------------------------------------------------- commands
def cmd_add(args):
    tickets = load_tickets()
    ticket = {
        "id": next_id(tickets),
        "title": args.title,
        "type": args.type,
        "priority": args.priority,
        "status": "To Do",
        "assignee": args.assignee or "Unassigned",
        "created": date.today().isoformat(),
        "due": args.due or "",
        "description": args.description or "",
    }
    tickets.append(ticket)
    save_tickets(tickets)
    print(f"Created {ticket['id']}: {ticket['title']}")


def cmd_list(args):
    tickets = load_tickets()
    if args.status:
        tickets = [t for t in tickets if t["status"].lower() == args.status.lower()]
    if args.assignee:
        tickets = [t for t in tickets if t["assignee"].lower() == args.assignee.lower()]
    if not tickets:
        print("No tickets found.")
        return
    print(f"{'ID':<9}{'Priority':<10}{'Status':<13}{'Assignee':<12}{'Due':<12}Title")
    print("-" * 80)
    for t in tickets:
        flag = " (OVERDUE)" if is_overdue(t) else ""
        print(f"{t['id']:<9}{t['priority']:<10}{t['status']:<13}"
              f"{t['assignee']:<12}{t['due']:<12}{t['title']}{flag}")


def cmd_update(args):
    tickets = load_tickets()
    for t in tickets:
        if t["id"].upper() == args.id.upper():
            for field in ["title", "type", "priority", "status", "assignee", "due", "description"]:
                value = getattr(args, field)
                if value is not None:
                    t[field] = value
            save_tickets(tickets)
            print(f"Updated {t['id']}")
            return
    sys.exit(f"Ticket {args.id} not found.")


def cmd_report(args):
    from report import build_tracker  # imported here so other commands don't need openpyxl
    tickets = load_tickets()
    out = build_tracker(tickets, args.out, STATUSES, PRIORITIES, is_overdue)
    print(f"Tracker sheet written to {out} ({len(tickets)} tickets)")


# -------------------------------------------------------------------- cli
def main():
    p = argparse.ArgumentParser(description="Simple ticket tracker")
    sub = p.add_subparsers(dest="command", required=True)

    a = sub.add_parser("add", help="create a new ticket")
    a.add_argument("--title", required=True)
    a.add_argument("--type", choices=TYPES, default="Task")
    a.add_argument("--priority", choices=PRIORITIES, default="Medium")
    a.add_argument("--assignee")
    a.add_argument("--due", type=valid_date, help="YYYY-MM-DD")
    a.add_argument("--description")
    a.set_defaults(func=cmd_add)

    l = sub.add_parser("list", help="list tickets")
    l.add_argument("--status", choices=STATUSES)
    l.add_argument("--assignee")
    l.set_defaults(func=cmd_list)

    u = sub.add_parser("update", help="update a ticket")
    u.add_argument("id")
    u.add_argument("--title")
    u.add_argument("--type", choices=TYPES)
    u.add_argument("--priority", choices=PRIORITIES)
    u.add_argument("--status", choices=STATUSES)
    u.add_argument("--assignee")
    u.add_argument("--due", type=valid_date)
    u.add_argument("--description")
    u.set_defaults(func=cmd_update)

    r = sub.add_parser("report", help="generate the Excel tracker sheet")
    r.add_argument("--out", default="tracker.xlsx")
    r.set_defaults(func=cmd_report)

    args = p.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
