# NUPI Ticket Loader

A single-page web app (published as a claude.ai artifact) for the people who
create tickets in NupiPay's TMS (tms.nupipay.com, EspoCRM). It replaces
one-at-a-time ticket entry with a grid, paste-from-sheet, or an Excel import,
and always shows a review of exactly what will be created before anything is
sent.

## How it is run (important)

The claude.ai artifact viewer blocks every network call a page makes except
script loads from a few CDNs, so the page **cannot talk to TMS while it is
open inside claude.ai**. The artifact therefore acts as the distribution point:
it shows a banner with "Save the app to this computer", which hands over the
complete page as `NUPI Ticket Loader.html`. Opened from disk (or hosted on any
ordinary web server, ideally on the TMS server itself) the page calls TMS
directly; TMS answers cross-origin requests with `Access-Control-Allow-Origin: *`
and accepts the standard `Authorization: Basic` header.

`scripts/build_standalone.py` writes the same complete document to
`dist/NUPI Ticket Loader.html` for sharing without going through claude.ai.

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

- `ticket-loader.html` the whole app (published artifact source)
- `template/NUPI Ticket Import Template.xlsx` offline sample of the template
  with SAMPLE names, built by `scripts/make_template.py`
- `test/page.test.js` headless test (jsdom + a fake EspoCRM): sign-in, token
  use, fallbacks, name matching, TSV/CSV parsing, paste, Excel import,
  template generation, validation, review, duplicate check, creation,
  failure + retry, drafts, sign-out. `npm test`

## Not yet verified against live TMS

- That the account used can list `User` and `Project` (fallbacks exist).
- Exact ProjectTask field names in this TMS (the page reads them at runtime).
- That `POST ProjectTask` accepts `columnId` on create; if TMS ignores it,
  tickets land in the default column and a follow-up `PUT` would be needed.
- EspoCRM accepting `Authorization: Basic user:token` without the
  `Espo-Authorization-By-Token` header (the TMS CORS policy does not allow
  that header from a browser). If it does not, the page silently keeps using
  the password for the tab.
