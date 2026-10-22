# NUPI Ticket Loader

A single-page web app for the people who create tickets in NupiPay's TMS
(tms.nupipay.com, EspoCRM). It replaces one-at-a-time ticket entry with a grid,
paste-from-sheet, or an Excel import, and always shows a review of exactly what
will be created before anything is sent.

**Live site (GitHub Pages, branch `gh-pages`):** https://akashtripathi9.github.io/tms-ticket-loader/

The deployed page is `index.html` on the `gh-pages` branch: the artifact source
(`ticket-loader.html`) wrapped in a complete document by `scripts/build_standalone.py`.
The page calls TMS directly from the browser; TMS answers cross-origin requests
with `Access-Control-Allow-Origin: *` and accepts the standard `Authorization: Basic`
header, so no server sits in between.

## What it does

- **Sign in with TMS credentials.** The password is exchanged for an EspoCRM
  token via `GET /api/v1/App/user`; only the token is kept, in the browser tab
  (sessionStorage).
- **Live reference data.** After sign-in it loads active users, projects, the
  ticket field definitions (`Metadata?key=entityDefs.ProjectTask.fields`),
  field labels and enum options (`I18n`), and board columns per project
  (derived from recent tickets). If the User or Project lists are restricted
  for the account, both are harvested from recent tickets instead.
- **Grid editor.** Title, Project, Assign to, Status and Description are
  always shown; any other editable ProjectTask field found in TMS (priority,
  dates, hours, custom checkboxes) can be switched on from "Columns".
  Rows can be added, duplicated, deleted, filtered, multi-selected and
  bulk-assigned a project, person or status. Enter in a title moves to the
  next row. Unsent work is kept as a draft in the browser.
- **Paste from Excel or Google Sheets.** A first row of headings is
  recognised (synonyms such as "Assigned to", "Owner", "Developer" work);
  without headings the order is Title, Project, Assign to, Status,
  Description. People, projects and statuses are matched to TMS by name,
  including partial names; anything unresolved is flagged red on the row.
- **Excel template.** "Download template" builds an .xlsx in the browser with
  the real users, projects and statuses as dropdowns (data validation lists on
  a Lists sheet), two example rows and an Instructions sheet. "Import Excel /
  CSV" (or drag-and-drop) reads it back. A sample sheet with placeholder names
  is embedded in the page for people who want to prepare rows before signing in.
- **Validation before review.** Required fields, duplicate titles within the
  batch, statuses that do not exist on the chosen project's board, and
  required custom fields.
- **Review.** Summary tiles, a plain-language "what will happen" list,
  tickets grouped by project, and an optional check against TMS for tickets
  that already have the same title in the same project.
- **Create.** One `POST /api/v1/ProjectTask` per row with ids and names
  (`projectId/projectName`, `assignedUserId/assignedUserName`,
  `columnId/columnName`). Progress per row, links to each created ticket,
  retry for failed rows, results as CSV (download or copy).

## Files

- `gh-pages` branch: `index.html`, the deployed page
- `ticket-loader.html` the page source (artifact form, no document skeleton)
- `scripts/build_standalone.py` wraps the source into `dist/NUPI Ticket Loader.html`
- `scripts/make_template.py` builds the offline sample template with openpyxl
- `test/page.test.js` headless test (jsdom + a fake EspoCRM): sign-in, token
  use, sandbox vs standalone mode, fallbacks, name matching, TSV/CSV parsing,
  paste, Excel import, template generation, validation, review, duplicate
  check, creation, failure + retry, drafts, sign-out. `npm install && npm test`

## Deploying a change

1. Edit `ticket-loader.html`, run `npm test`.
2. `python3 scripts/build_standalone.py` and copy `dist/NUPI Ticket Loader.html`
   to `index.html` on the `gh-pages` branch.

## Not yet verified against live TMS

- That the account used can list `User` and `Project` (fallbacks exist).
- Exact ProjectTask field names in this TMS (the page reads them at runtime).
- That `POST ProjectTask` accepts `columnId` on create; if TMS ignores it,
  tickets land in the default column and a follow-up `PUT` would be needed.
- EspoCRM accepting `Authorization: Basic user:token` without the
  `Espo-Authorization-By-Token` header (the TMS CORS policy does not allow
  that header from a browser). If it does not, the page silently keeps using
  the password for the tab.
