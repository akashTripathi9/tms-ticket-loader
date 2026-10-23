"""Builds the offline sample 'NUPI Ticket Import Template.xlsx'.

Columns, choice lists and notes come from scripts/tms_projecttask_schema.json, a snapshot of the real
ProjectTask field definitions read from TMS on 2026-10-23. People, projects, stages and ticket types are
placeholders marked SAMPLE; the template downloaded from NUPI Ticket Loader after sign-in carries the live
lists. Run: python3 scripts/make_template.py
"""
import json
from datetime import date, timedelta
from pathlib import Path
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "template" / "NUPI Ticket Import Template.xlsx"
OUT.parent.mkdir(parents=True, exist_ok=True)
SCHEMA = json.loads((ROOT / "scripts" / "tms_projecttask_schema.json").read_text())

ACCENT, ACCENT_SOFT, FONT, MAXR = "0E7C6B", "3A9B8B", "Arial", 1000
TODAY = date.today()
DATES = [TODAY + timedelta(days=i) for i in range(-14, 366)]  # two weeks back to a year ahead

USERS = ["SAMPLE – Akash Tripathi", "SAMPLE – Vipul Sharma", "SAMPLE – Riya Mehta", "SAMPLE – Ankit Dubey"]
PROJECTS = ["SAMPLE – Tech Board", "SAMPLE – Nets Board", "SAMPLE – Product Tank"]
STAGES = ["To Do", "In Progress", "BE in Progress", "FE in Progress", "BE Done", "FE Done", "Integration Done", "QA in Progress", "QA Done", "Done", "Released", "Backlog"]
TICKET_TYPES = ["Task", "SAMPLE – Bug"]
TEAMS = ["SAMPLE – Mobile", "SAMPLE – Backend"]

# ---- column model, same rules as the page (ticket-loader.html) ----------------------------
SKIP = {"id", "name", "description", "deleted", "createdAt", "modifiedAt", "createdBy", "modifiedBy", "assignedUser", "project", "column", "owner", "group", "groupColor", "groupOrder", "boardId", "status", "order", "boardOrder", "timeline", "cTrialEnum", "cStageForAnalysis", "attachments", "parent"}
ALLOWED = {"varchar", "text", "wysiwyg", "enum", "date", "datetime", "int", "float", "currency", "bool", "url", "multiEnum", "link", "linkMultiple"}
RANK = {"link": 1, "date": 2, "datetime": 3, "int": 4, "float": 4, "currency": 4, "enum": 5, "linkMultiple": 6, "multiEnum": 7, "bool": 8, "varchar": 9, "url": 9, "text": 10, "wysiwyg": 10}
fields, links, labels, tips = SCHEMA["fields"], SCHEMA["links"], SCHEMA["labels"], SCHEMA.get("tooltips", {})

def nice(k):
    import re
    k = re.sub(r"^c(?=[A-Z])", "", k); k = re.sub(r"([A-Z])", r" \1", k); return k.strip().capitalize() if k.islower() else k.strip()

extra = []
for k, d in fields.items():
    t = d.get("type")
    if t not in ALLOWED or k in SKIP or d.get("readOnly") or d.get("disabled") or d.get("utility") or d.get("notStorable"):
        continue
    import re as _re
    m = _re.match(r"^(.*?)(Ids?|Names?|Currency)$", k)
    if m and fields.get(m.group(1), {}).get("type") in ("link", "linkMultiple", "linkParent", "currency"):
        continue  # attribute shadow of a link field, not a field of its own
    entity = None
    if t in ("link", "linkMultiple"):
        entity = links.get(k, {}).get("entity")
        if entity not in ("User", "Team"):
            continue
    opts = [o for o in (d.get("options") or []) if o not in ("", None)] if d.get("options") else None
    default = d.get("default") or ("Normal" if k == "priority" and opts and "Normal" in opts else "Task" if k == "type" and opts and "Task" in opts else None)
    extra.append({"key": k, "label": labels.get(k) or nice(k), "type": "text" if t == "wysiwyg" else "float" if t == "currency" else t, "required": bool(d.get("required")), "options": opts, "entity": entity, "tip": d.get("tooltipText") or tips.get(k, ""), "default": default})
is_start = lambda f: 1 if any(w in (f["key"] + " " + f["label"]).lower() for w in ("start", "begin", "from")) else 0
extra.sort(key=lambda f: (-int(f["required"]), RANK.get(f["type"], 8), -is_start(f), f["label"]))

CORE_HEAD = [
    {"key": "name", "label": "Ticket title", "required": True, "type": "varchar", "note": "Required. A short, specific title. Rows without a title are skipped."},
    {"key": "projectId", "label": "Project", "required": True, "type": "project", "note": "Required. Pick from the dropdown; the list comes from TMS."},
    {"key": "groupId", "label": labels.get("group", "Ticket Type"), "type": "group", "note": "Ticket type such as Task or Bug; blank = Task. Types are per project."},
    {"key": "assignedUserId", "label": "Assign to", "type": "user", "note": "Optional. Pick a person, or leave blank to create the ticket unassigned."},
    {"key": "columnId", "label": labels.get("column", "Stage"), "type": "status", "note": "Optional. Board stage the ticket starts in; blank = the project's first stage (To Do)."},
]
CORE_TAIL = [{"key": "description", "label": "Description", "type": "text", "note": "Recommended. What needs to be done and how we will know it is done."}]

def note_for(f):
    if f.get("note"):
        return f["note"]
    t = f["type"]; base = (f["tip"] + " ") if f.get("tip") else ""
    how = {"date": "Date: click the arrow to pick, or type YYYY-MM-DD / DD-MM-YYYY.", "datetime": "Date and time, e.g. 2026-10-23 14:30.", "bool": "Yes or No.", "int": "Whole number.", "float": "Number, e.g. 4 or 2.5.", "url": "A full link starting with http(s)://.", "text": "Free text.", "varchar": "Free text."}.get(t)
    if t == "enum": how = "Pick one: " + ", ".join(f["options"]) + "."
    elif t == "multiEnum": how = "One or more of: " + ", ".join(f["options"]) + " (comma separated)."
    elif t == "link": how = f"Pick one {f['entity'].lower()} from the dropdown."
    elif t == "linkMultiple": how = f"One or more {f['entity'].lower()}s, comma separated."
    req = ("Required; blank = %s. " % f["default"]) if f["required"] and f.get("default") else "Required. " if f["required"] else "Optional. "
    return req + base + how

COLUMNS = CORE_HEAD + extra + CORE_TAIL
def width(f):
    return {"name": 46, "description": 64}.get(f["key"]) or {"date": 14, "datetime": 20, "int": 12, "float": 14, "bool": 12, "text": 40}.get(f["type"], 22)

wb = Workbook(); ws = wb.active; ws.title = "Tickets"; lists = wb.create_sheet("Lists"); info = wb.create_sheet("Instructions")

# ---- Lists sheet --------------------------------------------------------------------------
LISTS = [("Users", USERS), ("Projects", PROJECTS), ("Stages", STAGES), ("Ticket types", TICKET_TYPES), ("Teams", TEAMS)]
for f in extra:
    if f["type"] in ("enum", "multiEnum") and f["options"]:
        LISTS.append((f["label"], f["options"]))
LISTS.append(("Dates", DATES))
for ci, (title, values) in enumerate(LISTS, start=1):
    c = lists.cell(row=1, column=ci, value=title); c.font = Font(name=FONT, bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=ACCENT)
    for r, v in enumerate(values, start=2):
        cell = lists.cell(row=r, column=ci, value=v); cell.font = Font(name=FONT)
        if isinstance(v, date): cell.number_format = "yyyy-mm-dd"
    lists.column_dimensions[get_column_letter(ci)].width = 14 if title == "Dates" else 30
lists.freeze_panes = "A2"
lists.cell(row=2, column=len(LISTS) + 2, value="Users, Projects, Stages, Ticket types and Teams are SAMPLE values; the other lists are the real choices from TMS (as of %s). The template downloaded from NUPI Ticket Loader after sign-in carries the real names for everything." % SCHEMA["capturedAt"]).font = Font(name=FONT, italic=True, color="6B7A72")
def list_ref(title):
    i = next(i for i, (t, _) in enumerate(LISTS, start=1) if t == title); L = get_column_letter(i); n = len(LISTS[i - 1][1])
    return f"Lists!${L}$2:${L}${max(2, n + 1)}"

# ---- Tickets sheet ------------------------------------------------------------------------
thin = Side(style="thin", color="D8DED6")
for ci, f in enumerate(COLUMNS, start=1):
    c = ws.cell(row=1, column=ci, value=f["label"] + (" *" if f.get("required") else ""))
    c.font = Font(name=FONT, bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=ACCENT if f.get("required") else ACCENT_SOFT); c.alignment = Alignment(vertical="center")
    c.comment = Comment(note_for(f), "NUPI Ticket Loader"); c.comment.width, c.comment.height = 320, 110
    ws.column_dimensions[get_column_letter(ci)].width = width(f)
ws.row_dimensions[1].height = 26; ws.freeze_panes = "C2"; ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}1"

def add_dv(f, ci, **kw):
    L = get_column_letter(ci)
    dv = DataValidation(allow_blank=not f.get("required"), showInputMessage=True, promptTitle=f["label"][:32], **kw)
    dv.add(f"{L}2:{L}{MAXR}"); ws.add_data_validation(dv)

for ci, f in enumerate(COLUMNS, start=1):
    t, k = f["type"], f["key"]
    if k == "projectId": add_dv(f, ci, type="list", formula1=list_ref("Projects"), showErrorMessage=True, errorTitle="Unknown project", error="Pick a project from the dropdown.", prompt="Required. Pick the project.")
    elif k == "assignedUserId": add_dv(f, ci, type="list", formula1=list_ref("Users"), showErrorMessage=True, errorTitle="Unknown user", error="Pick a person from the dropdown, or leave blank.", prompt="Pick the person, or leave blank for unassigned.")
    elif k == "columnId": add_dv(f, ci, type="list", formula1=list_ref("Stages"), showErrorMessage=True, errorTitle="Unknown stage", error="Pick a stage, or leave blank for the first stage.", prompt="Board stage to start in. Blank = To Do.")
    elif k == "groupId": add_dv(f, ci, type="list", formula1=list_ref("Ticket types"), showErrorMessage=True, errorTitle="Unknown ticket type", error="Pick a ticket type, or leave blank for Task.", prompt="Task or Bug. Blank = Task.")
    elif t in ("link", "linkMultiple") and f["entity"] == "User": add_dv(f, ci, type="list", formula1=list_ref("Users"), showErrorMessage=(t == "link"), prompt="Pick a person." if t == "link" else "Pick a person; for several, type names separated by commas.")
    elif t in ("link", "linkMultiple") and f["entity"] == "Team": add_dv(f, ci, type="list", formula1=list_ref("Teams"), showErrorMessage=(t == "link"), prompt="Pick a team." if t == "link" else "Pick a team; for several, type names separated by commas.")
    elif t in ("enum", "multiEnum") and f["options"]: add_dv(f, ci, type="list", formula1=list_ref(f["label"]), showErrorMessage=(t == "enum"), errorTitle="Not a valid choice", error="Pick one of the listed values.", prompt="Pick one value." if t == "enum" else "Pick a value; for several, type them separated by commas.")
    elif t in ("date", "datetime"):
        # A plain .xlsx cannot host a calendar control, so each date cell gets a click-to-pick list of dates; typing still works.
        add_dv(f, ci, type="list", formula1=list_ref("Dates"), showErrorMessage=False, prompt="Click the arrow and pick the date, or type one (2026-10-23 or 23/10/2026). In Google Sheets, double-click the cell for a calendar.")
        for r in range(2, MAXR + 1): ws.cell(row=r, column=ci).number_format = "yyyy-mm-dd" if t == "date" else "yyyy-mm-dd hh:mm"
    elif t == "int": add_dv(f, ci, type="whole", operator="greaterThanOrEqual", formula1="0", showErrorMessage=True, error="Enter a whole number.", prompt="Whole number.")
    elif t == "float": add_dv(f, ci, type="decimal", operator="greaterThanOrEqual", formula1="0", showErrorMessage=True, error="Enter a number, e.g. 4 or 2.5.", prompt="Number, e.g. 4 or 2.5.")
    elif t == "bool": add_dv(f, ci, type="list", formula1='"Yes,No"', showErrorMessage=True, error="Yes or No.", prompt="Yes or No.")

# example rows (grey italics) -----------------------------------------------------------------
def example(f, which):
    k, t = f["key"], f["type"]
    fixed = {"name": ["EXAMPLE – Add OTP resend cooldown on login screen", "EXAMPLE – Fix settlement report export timeout"], "projectId": [PROJECTS[0], PROJECTS[1]], "groupId": ["Task", "SAMPLE – Bug"], "assignedUserId": [USERS[1], ""], "columnId": ["To Do", ""],
             "description": ["Show a 30 second countdown after the first OTP, disable the resend button meanwhile, and log the resend event to analytics.", "Export of more than 5000 rows times out. Move the export to a background job and email the file when ready."]}
    if k in fixed: return fixed[k][which]
    if which == 1 and not f.get("default"): return None
    if f.get("default"): return f["default"]
    if t == "enum": return f["options"][min(1, len(f["options"]) - 1)] if f["options"] else None
    if t == "multiEnum": return f["options"][0] if f["options"] else None
    if t == "date": return DATES[14] if is_start(f) else DATES[18]
    if t == "int": return 3
    if t == "float": return 6
    if t == "bool": return "No"
    if t in ("link", "linkMultiple"): return USERS[0] if f["entity"] == "User" else TEAMS[0]
    if t == "url": return "https://www.figma.com/file/EXAMPLE"
    if t == "varchar": return "NETS-1042"
    return None
for r in (2, 3):
    for ci, f in enumerate(COLUMNS, start=1):
        v = example(f, r - 2)
        c = ws.cell(row=r, column=ci, value=v); c.font = Font(name=FONT, italic=True, color="6B7A72"); c.alignment = Alignment(wrap_text=True, vertical="top"); c.border = Border(bottom=thin)
        if isinstance(v, date): c.number_format = "yyyy-mm-dd"
for r in range(4, 60):
    for ci in range(1, len(COLUMNS) + 1):
        c = ws.cell(row=r, column=ci); c.font = Font(name=FONT); c.alignment = Alignment(wrap_text=True, vertical="top")

# ---- Instructions sheet -------------------------------------------------------------------
info.column_dimensions["A"].width = 118
required_names = ", ".join(f["label"] for f in COLUMNS if f.get("required") and not f.get("default"))
LINES = [
    ("How to use this sheet", 14, True),
    ("1. Fill one ticket per row on the “Tickets” sheet. Delete the two grey EXAMPLE rows before importing.", 11, False),
    ("2. Columns marked * must be filled: %s. Priority is also required by TMS but blank means Normal. Every other column is optional and matches a field on the TMS ticket form; leave what you do not need empty." % required_names, 11, False),
    ("3. Project, Ticket type, Assign to, Stage, Tester, Teams and every choice column are dropdowns fed by the “Lists” sheet. In this sample the people, projects, stages and ticket types are placeholders; the template downloaded from NUPI Ticket Loader after sign-in carries the real names from TMS.", 11, False),
    ("4. Date columns: click the cell, then the arrow, and pick the date from the list (two weeks back to a year ahead). Typing works too: 2026-10-23 or 23/10/2026. In Google Sheets, double-click a date cell to get a calendar.", 11, False),
    ("5. Leave Stage blank to start in the project’s first stage (To Do), Ticket type blank for Task, Priority blank for Normal, Assign to blank for unassigned. The person importing the sheet is recorded as the ticket’s Creator.", 11, False),
    ("6. In NUPI Ticket Loader, click “Import Excel / CSV” and choose this file. Check the preview, fix anything marked red, click Review, then Confirm. Nothing is created in TMS until you confirm.", 11, False),
    ("", 11, False),
    ("Legend", 12, True),
    ("Dark green header = required column. Light green header = optional. Grey italic rows = examples to delete.", 11, False),
    ("", 11, False),
    ("Column guide (from the TMS ticket form, %s)" % SCHEMA["capturedAt"], 12, True),
] + [("%s%s — %s" % (f["label"], " (required)" if f.get("required") else "", note_for(f)), 11, False) for f in COLUMNS]
for r, (text, size, bold) in enumerate(LINES, start=1):
    c = info.cell(row=r, column=1, value=text); c.font = Font(name=FONT, size=size, bold=bold); c.alignment = Alignment(wrap_text=True, vertical="top")

wb.save(OUT)
print(OUT, "| columns:", len(COLUMNS), "| lists:", [t for t, _ in LISTS])
