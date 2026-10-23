# NUPI Ticket Loader

A single-page web app for the people who
create tickets in NupiPay's TMS (tms.nupipay.com, EspoCRM). It replaces
one-at-a-time ticket entry with a grid, paste-from-sheet, or an Excel import,
and always shows a review of exactly what will be created before anything is
sent.

## Where it runs

The app is a static page served from GitHub Pages at
https://akashtripathi9.github.io/tms-ticket-loader/ (branch `gh-pages`,
repo `akashTripathi9/tms-ticket-loader`). It talks to TMS directly from the
browser: TMS answers cross-origin requests with `Access-Control-Allow-Origin: *`
and accepts the standard `Authorization: Basic` header. There is no server in
between and nothing to install. `scripts/build_standalone.py` wraps the page
source into `dist/NUPI Ticket Loader.html`, which is what gets deployed as
`site/index.html`.

## What it does

- **Sign in with TMS credentials.** The password is exchanged for an EspoCRM
  token via `GET /api/v1/App/user`; only the token is kept, in the browser
  tab (sessionStorage). TMS already sends CORS headers for successful
  responses, so the page talks to the API directly; no server in between.
- **Live reference data.** After sign-in it loads active users, projects, the
  ticket field definitions (`Metadata?key=entityDefs.ProjectTask.fields`),
  field labels and enum options (`I18n`), and board columns per project
  (derived from recent tickets). If the User or Project lists are restricted
  for the account, both are harvested from recent tickets instead.
- **Grid editor.** Title, Project, Assign to, Status and Description are
  always shown, together with every field TMS marks as required. Any other
  editable ProjectTask field found in TMS (priority, dates, hours, reviewer,
  teams, custom checkboxes…) can be switched on from "Columns" ("Show all").
  Link fields that point at User or Team become dropdowns; date fields use
  the browser's calendar picker.
  Rows can be added, duplicated, deleted, filtered, multi-selected and
  bulk-assigned a project, person or status. Enter in a title moves to the
  next row. Unsent work is kept as a draft in the browser.
- **Paste from Excel or Google Sheets.** A first row of headings is
  recognised (synonyms such as "Assigned to", "Owner", "Developer" work);
  without headings the order is Title, Project, Assign to, Status,
  Description. People, projects and statuses are matched to TMS by name,
  including partial names; anything unresolved is flagged red on the row.
- **Excel template.** "Download template" builds an .xlsx in the browser with
  **every** field of the TMS ticket form as a column (required ones first,
  dark header, `*`), the real users, projects, statuses, teams and choice
  lists as dropdowns (data validation lists on a Lists sheet), a note on each
  heading, two example rows, and an Instructions sheet with a column guide.
  Date columns get a click-to-pick list of dates (two weeks back to a year
  ahead) because a plain .xlsx cannot carry a calendar control; typing a date
  still works and Google Sheets shows a calendar on double-click. "Import
  Excel / CSV" (or drag-and-drop) reads it back and understands dates in ISO,
  DD/MM/YYYY, DD-MM-YYYY, "23 Oct 2026" and Excel serial form.
- **Validation before review.** Required fields (including ones TMS marks
  required), duplicate titles within the batch, statuses that do not exist on
  the chosen project's board, unreadable dates or numbers, unknown names, and
  a due date earlier than the start date.
- **Review.** Summary tiles, a plain-language "what will happen" list,
  tickets grouped by project, and an optional check against TMS for tickets
  that already have the same title in the same project.
- **Create.** One `POST /api/v1/ProjectTask` per row with ids and names
  (`projectId/projectName`, `assignedUserId/assignedUserName`,
  `columnId/columnName`, `<link>Id/<link>Name`, `<multi>Ids/<multi>Names`),
  date-times converted to UTC `YYYY-MM-DD HH:MM:SS`. Progress per row, links to each created ticket,
  retry for failed rows, results as CSV (download or copy).

## Files

- `ticket-loader.html` the whole app (page source without the document wrapper)
- `template/NUPI Ticket Import Template.xlsx` offline sample of the template
  with SAMPLE names, built by `scripts/make_template.py`
- `test/page.test.js` headless test (jsdom + a fake EspoCRM): sign-in, token
  use, fallbacks, name matching, TSV/CSV parsing, paste, Excel import,
  template generation, validation, review, duplicate check, creation,
  failure + retry, drafts, sign-out. `npm test`

## What a ticket needs (verified against TMS on 2026-10-23)

`scripts/tms_projecttask_schema.json` is a snapshot of the real ProjectTask
field definitions, links and labels, read from TMS with a regular user's
session. The page reads the same thing live after sign-in; the snapshot drives
the offline sample sheet and documents the form.

Required by TMS on create:

| Field (TMS label)   | How the loader fills it                                   |
|---------------------|-----------------------------------------------------------|
| name                | Ticket title column                                       |
| project             | Project column (dropdown)                                 |
| group "Ticket Type" | Ticket Type column; blank = the project's "Task" group    |
| column "Stage"      | Stage column; blank = the project's first stage (To Do)   |
| owner "Creator"     | always the signed-in user                                 |
| priority            | Priority column; blank = Normal                           |

`boardId` is sent alongside the stage when known (from ProjectColumn or recent
tickets), as the TMS UI does. `status` is read-only ("Ignore") and left to TMS.

Optional fields found on the form and offered as columns: Type (Task /
Milestone), Date Start, Date End, Planned Start/End Date, Due Date, Target
Date, Estimate Hours, Hours, Points, Client Name, Epic Name, Category, Product
Family / Type / Product, Feature Type, Features, Environment, Health, Impact,
Timeline Confidence, Paid CR, Ticket Sub-Type, Issue Type / Category / Sub
Category / Sub Type, Management Response, Allocation (multi), Tester (users,
multi), Reference Ticket Number, Prototype Link, Working on it today.
Read-only, computed and attachment fields are left out.

## Not yet verified against live TMS

- That the account used can list `User`, `Project`, `Team`, `ProjectGroup`
  and `ProjectColumn` (fallbacks harvest all of them from recent tickets).
- An actual `POST ProjectTask` from the page (the payload mirrors the one the
  TMS UI sends, including `groupId`, `ownerId`, `columnId` and `boardId`).
- EspoCRM accepting `Authorization: Basic user:token` without the
  `Espo-Authorization-By-Token` header (the TMS CORS policy does not allow
  that header from a browser). If it does not, the page silently keeps using
  the password for the tab.
