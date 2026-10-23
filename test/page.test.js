/* Headless test of NUPI Ticket Loader: loads the page in jsdom with a fake TMS
   (EspoCRM-shaped responses), signs in, imports, pastes, validates, reviews and
   creates. Fails on any JS error. Run: npm test */
const fs = require("fs");
const path = require("path");
const assert = require("assert");
const { JSDOM, VirtualConsole } = require("jsdom");
const ExcelJS = require("exceljs");
const EXCELJS_SRC = fs.readFileSync(require.resolve("exceljs/dist/exceljs.min.js"), "utf8");

const html = fs.readFileSync(path.join(__dirname, "..", "ticket-loader.html"), "utf8").replace(/<script src="https:[^"]+"><\/script>/g, "");

// ---------- fake TMS ----------
const USERS = [{ id: "u1", name: "Akash Tripathi", userName: "akash", type: "admin", isActive: true }, { id: "u2", name: "Vipul Sharma", userName: "vipul", type: "regular", isActive: true }, { id: "u3", name: "Riya Mehta", userName: "riya", type: "regular", isActive: true }, { id: "u4", name: "Portal Bot", userName: "bot", type: "api", isActive: true }];
const PROJECTS = [{ id: "p1", name: "NUPI App" }, { id: "p2", name: "Backoffice" }];
const TASKS = [
  { id: "t1", name: "Existing ticket one", projectId: "p1", projectName: "NUPI App", boardId: "b1", columnId: "c1", columnName: "To Do", groupId: "g1", groupName: "Task", assignedUserId: "u1", assignedUserName: "Akash Tripathi" },
  { id: "t2", name: "Old bug", projectId: "p1", projectName: "NUPI App", boardId: "b1", columnId: "c2", columnName: "In Progress", groupId: "g2", groupName: "Bug", assignedUserId: "u2", assignedUserName: "Vipul Sharma" },
  { id: "t3", name: "Report export", projectId: "p2", projectName: "Backoffice", boardId: "b2", columnId: "c9", columnName: "Backlog", groupId: "g3", groupName: "Task", assignedUserId: "u3", assignedUserName: "Riya Mehta" },
];
const GROUPS = [{ id: "g1", name: "Task", projectId: "p1", order: 1 }, { id: "g2", name: "Bug", projectId: "p1", order: 2 }, { id: "g3", name: "Task", projectId: "p2", order: 1 }];
const COLUMNS = [{ id: "c2", name: "In Progress", boardId: "b1", order: 2 }, { id: "c1", name: "To Do", boardId: "b1", order: 1 }, { id: "c3", name: "Done", boardId: "b1", order: 3 }, { id: "c9", name: "Backlog", boardId: "b2", order: 1 }];
const DEFS = { name: { type: "varchar", required: true }, description: { type: "text" }, project: { type: "link", required: true }, column: { type: "link", required: true }, group: { type: "link", required: true }, owner: { type: "link", required: true }, assignedUser: { type: "link" },
  priority: { type: "enum", required: true, options: ["Low", "Normal", "High", "Urgent"] }, status: { type: "enum", readOnly: true, options: ["Not Started", "Started"] }, type: { type: "enum", options: ["Task", "Milestone"] },
  dateStart: { type: "date" }, dateEnd: { type: "date" }, cPlannedStartDate: { type: "date" }, cPlannedEndDate: { type: "date" }, cDueDate: { type: "date" }, cExpectedClosureDateKAM: { type: "date" }, cEstimateHours: { type: "float" }, hours: { type: "float" }, points: { type: "int" },
  cClientName: { type: "enum", options: ["", "STIHL", "NUPI APP"] }, cResourceAllocation: { type: "multiEnum", options: ["", "Akash", "Vipul"] }, cUsers: { type: "linkMultiple" }, teams: { type: "linkMultiple" }, cWorkingOnItToday: { type: "bool" }, cPrototypeLink: { type: "url" }, cReferenceTicketNumber: { type: "varchar" }, cReminderAt: { type: "datetime" },
  boardId: { type: "foreign", utility: true }, groupColor: { type: "foreign" }, cStageForAnalysis: { type: "enum", readOnly: true, options: ["To Do"] }, attachments: { type: "attachmentMultiple" }, createdAt: { type: "datetime", readOnly: true }, cTotalLoggedHours: { type: "float", readOnly: true }, fromTasks: { type: "linkMultiple" }, parentTask: { type: "link" }, timeline: { type: "base", notStorable: true } };
const LINKS = { assignedUser: { type: "belongsTo", entity: "User" }, project: { type: "belongsTo", entity: "Project" }, column: { type: "belongsTo", entity: "ProjectColumn" }, group: { type: "belongsTo", entity: "ProjectGroup" }, owner: { type: "belongsTo", entity: "User" }, cUsers: { type: "hasMany", entity: "User" }, teams: { type: "hasMany", entity: "Team" }, fromTasks: { type: "hasMany", entity: "ProjectTask" }, parentTask: { type: "belongsTo", entity: "ProjectTask" } };
const TEAMS = [{ id: "tm1", name: "Mobile" }, { id: "tm2", name: "Backend" }];
const I18N = { ProjectTask: { fields: { priority: "Priority", dateStart: "Date Start", dateEnd: "Date End", cPlannedStartDate: "Planned Start Date", cPlannedEndDate: "Planned End Date", cDueDate: "Due Date", cExpectedClosureDateKAM: "Target Date", cEstimateHours: "Estimate Hours", hours: "Hours", points: "Points", cClientName: "Client Name", cResourceAllocation: "Allocation", cUsers: "Tester", teams: "Teams", cWorkingOnItToday: "Working on it today", cPrototypeLink: "Prototype Link", cReferenceTicketNumber: "Reference Ticket Number", cReminderAt: "Reminder", column: "Stage", group: "Ticket Type", owner: "Creator", type: "Type" }, options: { priority: { Low: "Low", Normal: "Normal", High: "High", Urgent: "Urgent" } }, tooltips: { dateStart: "Ideation Phase has started" } } };

function makeFetch(opts) {
  opts = opts || {};
  const calls = []; const created = [];
  const fetchImpl = async (url, init) => {
    init = init || {};
    if (!/^https?:/.test(url)) { calls.push({ method: "GET", path: url }); return { ok: url.endsWith(".xlsx"), status: url.endsWith(".xlsx") ? 200 : 404, headers: { get: () => null }, blob: async () => ({ size: 1234, type: "application/octet-stream" }), text: async () => "" }; }
    const u = new URL(url); const p = u.pathname.replace(/^\/api\/v1\//, ""); const method = init.method || "GET";
    const auth = (init.headers || {}).Authorization || "";
    calls.push({ method, path: p, query: Object.fromEntries(u.searchParams.entries()), body: init.body ? JSON.parse(init.body) : undefined, auth });
    const json = (status, body) => ({ ok: status < 400, status, headers: { get: () => null }, text: async () => (body === undefined ? "" : JSON.stringify(body)) });
    const decoded = Buffer.from(auth.replace(/^Basic /, ""), "base64").toString("utf8");
    const [user, secret] = decoded.split(":");
    if (user !== "akash" || !(secret === "pw" || secret === "TOKEN123")) { const e = new TypeError("Failed to fetch"); throw e; } // TMS: no CORS headers on 401
    if (p === "App/user") return json(200, { user: { id: "u1", name: "Akash Tripathi", isAdmin: true }, token: "TOKEN123" });
    if (p === "User") { if (opts.noUsers) return json(403, {}); return json(200, { list: USERS }); }
    if (p === "Project") { if (opts.noProjects) return json(403, {}); return json(200, { list: PROJECTS }); }
    if (p === "Metadata") return json(200, u.searchParams.get("key") === "entityDefs.ProjectTask.links" ? LINKS : DEFS);
    if (p === "Team") return json(200, { list: TEAMS });
    if (p === "ProjectGroup") { if (opts.noGroups) return json(403, {}); return json(200, { list: GROUPS }); }
    if (p === "ProjectColumn") { if (opts.noColumns) return json(403, {}); return json(200, { list: COLUMNS }); }
    if (p === "I18n") return json(200, I18N);
    if (p === "ProjectTask" && method === "GET") {
      const q = Object.fromEntries(u.searchParams.entries());
      let list = TASKS.concat(created);
      if (q["where[0][attribute]"] === "projectId") list = list.filter((t) => t.projectId === q["where[0][value]"]);
      if (q["where[1][attribute]"] === "name") { const names = Object.entries(q).filter(([k]) => k.startsWith("where[1][value]")).map(([, v]) => v.toLowerCase()); list = list.filter((t) => names.includes(t.name.toLowerCase())); }
      return json(200, { list });
    }
    if (p === "ProjectTask" && method === "POST") {
      const b = JSON.parse(init.body);
      if (opts.failName && b.name === opts.failName && !opts.failOnce--) return json(403, { message: "forbidden" });
      const t = Object.assign({ id: "n" + (created.length + 1), number: 500 + created.length }, b); created.push(t); return json(200, t);
    }
    if (p === "App/action/destroyAuthToken") return json(200, true);
    return json(404, {});
  };
  return { fetchImpl, calls, created };
}

function boot(opts) {
  opts = opts || {};
  const errors = [];
  const vc = new VirtualConsole();
  vc.on("jsdomError", (e) => errors.push(String(e.detail || e.message || e)));
  const fake = makeFetch(opts);
  const dom = new JSDOM("<!doctype html><html><head><meta charset='utf-8'></head><body>" + html + "</body></html>", {
    runScripts: "outside-only", pretendToBeVisual: true, virtualConsole: vc, url: "https://claude.ai/artifact/test",
    beforeParse(window) {
      window.fetch = fake.fetchImpl; window.scrollTo = () => {};
      window.Element.prototype.scrollIntoView = function () {};
      if (opts.sandbox) window.claude = { use: async (name) => (name === "downloads" ? { save: async (req) => { window.__saved = req; return { status: "saved" }; } } : null) };
      // Standalone mode saves through a blob link; record it instead of navigating.
      window.URL.createObjectURL = (blob) => { window.__blob = blob; return "blob:test"; }; window.URL.revokeObjectURL = () => {};
      window.HTMLAnchorElement.prototype.click = function () { if (this.download) window.__saved = { filename: this.download, data: window.__blob }; };
    },
  });
  const { window } = dom;
  // Run the real browser build of ExcelJS inside the page realm, as a browser would (arrays/dates must share the realm).
  window.eval(EXCELJS_SRC);
  for (const s of [...window.document.querySelectorAll("script")]) { try { window.eval(s.textContent); } catch (e) { errors.push("script eval: " + e.stack); } }
  return { window, document: window.document, errors, fake, $: (id) => window.document.getElementById(id), TL: window.TL };
}
const tick = (ms) => new Promise((r) => setTimeout(r, ms || 0));
async function settle(w, n) { for (let i = 0; i < (n || 20); i++) await tick(2); }
async function login(w) {
  w.$("user").value = "akash"; w.$("pass").value = "pw";
  w.$("loginForm").dispatchEvent(new w.window.Event("submit", { cancelable: true }));
  await settle(w, 60);
}
const setVal = (w, el, v) => { el.value = v; el.dispatchEvent(new w.window.Event("input", { bubbles: true })); el.dispatchEvent(new w.window.Event("change", { bubbles: true })); };

const eq = (a, b, m) => assert.strictEqual(JSON.stringify(a), JSON.stringify(b), m);
let failures = 0, passes = 0;
function t(name, fn) { return Promise.resolve().then(fn).then(() => { passes++; console.log("  ok  ", name); }, (e) => { failures++; console.log("  FAIL", name, "\n       ", (e && e.message || String(e)).split("\n").slice(0, 8).join("\n        ")); }); }

(async () => {
  console.log("Sign-in and reference data");
  {
    const w = boot();
    await t("boots without JS errors, shows sign-in", async () => { assert.deepStrictEqual(w.errors, []); assert.ok(!w.$("signin").hidden); assert.ok(w.$("edit").hidden); });
    await t("wrong password shows a clear error", async () => { w.$("user").value = "akash"; w.$("pass").value = "nope"; w.$("loginForm").dispatchEvent(new w.window.Event("submit", { cancelable: true })); await settle(w, 30); assert.match(w.$("loginErr").textContent, /username or password was refused/); });
    await t("login exchanges password for token and uses it afterwards", async () => { await login(w); assert.ok(!w.$("edit").hidden); const later = w.fake.calls.filter((c) => c.path !== "App/user"); assert.ok(later.length > 3); assert.ok(later.every((c) => Buffer.from(c.auth.slice(6), "base64").toString().endsWith(":TOKEN123"))); assert.strictEqual(w.$("pass").value, ""); });
    await t("users exclude api accounts; projects, columns and extra fields loaded", async () => { eq(w.TL.ref.users.map((u) => u.id), ["u1", "u2", "u3"]); assert.strictEqual(w.TL.ref.projects.length, 2); eq(w.TL.ref.columnsByProject.p1.map((c) => c.name), ["To Do", "In Progress", "Done"]); eq(w.TL.ref.fields.map((f) => f.key), ["priority", "dateStart", "cPlannedStartDate", "dateEnd", "cDueDate", "cPlannedEndDate", "cExpectedClosureDateKAM", "cReminderAt", "cEstimateHours", "hours", "points", "cClientName", "type", "teams", "cUsers", "cResourceAllocation", "cWorkingOnItToday", "cPrototypeLink", "cReferenceTicketNumber"]); assert.strictEqual(w.TL.ref.fields[0].label, "Priority"); assert.strictEqual(w.TL.ref.fields.find((f) => f.key === "cUsers").entity, "User"); eq(w.TL.ref.teams.map((t) => t.name), ["Mobile", "Backend"]); eq(w.TL.ref.fields.find((f) => f.key === "cClientName").options, ["STIHL", "NUPI APP"], "empty enum option dropped"); assert.strictEqual(w.TL.ref.fields.find((f) => f.key === "dateStart").tooltip, "Ideation Phase has started"); });
    await t("required-with-default fields (Priority) are shown and prefilled; stages and ticket types come per project with boards", async () => { assert.ok(w.TL.visibleCols.has("priority")); assert.ok([...w.$("headRow").children].some((th) => th.textContent.includes("Priority"))); eq(w.TL.ref.defaults, { priority: "Normal", type: "Task" }); assert.strictEqual(w.TL.rows[0].priority, "Normal"); eq(w.TL.ref.columnsByProject.p1.map((c) => c.name), ["To Do", "In Progress", "Done"], "full board order from ProjectColumn"); eq(w.TL.groupsFor("p1").map((g) => g.name), ["Task", "Bug"]); assert.strictEqual(w.TL.defaultGroup("p2").id, "g3"); assert.strictEqual(w.TL.defaultColumn("p1").id, "c1"); assert.strictEqual(w.TL.ref.boardByColumn.c3, "b1"); assert.ok(w.TL.ref.groupRequired && w.TL.ref.columnRequired); const gl = [...w.$("headRow").children].map((th) => th.textContent.trim()); eq(gl.slice(2, 7), ["Ticket title *", "Project *", "Ticket Type", "Assign to", "Stage"]); });
    await t("dates read from any common format, day first when ambiguous", async () => { const P = w.TL.parseDateText; eq([P("2026-10-23"), P("23/10/2026"), P("23-10-2026"), P("23.10.26"), P("23 Oct 2026"), P("Oct 23, 2026"), P("3rd Nov 2026"), P(46000), P(new w.window.Date(Date.UTC(2026, 0, 5))), P("31/02/2026"), P("hello"), P("")], ["2026-10-23", "2026-10-23", "2026-10-23", "2026-10-23", "2026-10-23", "2026-10-23", "2026-11-03", "2025-12-09", "2026-01-05", null, null, null]); assert.strictEqual(w.TL.toEspoDateTime("2026-10-23T14:30"), new w.window.Date("2026-10-23T14:30").toISOString().slice(0, 19).replace("T", " ")); });
    await t("headings map to runtime fields: Due / Estimate / Team / Reviewer", async () => { const H = w.TL.headerToKey; eq([H("Due"), H("Deadline"), H("End date"), H("Start"), H("Planned start"), H("Target date"), H("Estimate"), H("Hours"), H("Teams"), H("Tester"), H("Ticket Type"), H("Stage"), H("Type"), H("Client"), H("Allocation"), H("Priority *")], ["cDueDate", "cDueDate", "dateEnd", "dateStart", "cPlannedStartDate", "cExpectedClosureDateKAM", "cEstimateHours", "hours", "teams", "cUsers", "groupId", "columnId", "type", "cClientName", "cResourceAllocation", "priority"]); });
    await t("grid renders one empty row with live dropdowns", async () => { const sel = w.document.querySelector('select[data-k="projectId"]'); assert.ok(sel); assert.strictEqual(sel.options.length, 3); assert.strictEqual(w.document.querySelector('select[data-k="assignedUserId"]').options.length, 4); assert.ok(w.$("who").textContent.includes("Akash Tripathi")); });
    await t("choosing a project narrows the status list", async () => { const tr = w.document.querySelector("#rows tr"); setVal(w, tr.querySelector('select[data-k="projectId"]'), "p2"); const st = tr.querySelector('select[data-k="columnId"]'); eq([...st.options].map((o) => o.textContent), ["Default: Backlog", "Backlog"]); });
    await t("columns popover lists extra TMS fields; enabling one adds a column", async () => { w.$("colsBtn").click(); const cb = w.$("colsPop").querySelector('input[data-k="priority"]'); assert.ok(cb); cb.checked = true; cb.dispatchEvent(new w.window.Event("change", { bubbles: true })); assert.ok([...w.$("headRow").children].some((th) => th.textContent.includes("Priority"))); assert.strictEqual(w.document.querySelector('select[data-k="priority"]').options.length, 5); w.$("colsAll").click(); assert.ok(w.TL.visibleCols.has("teams") && w.TL.visibleCols.has("cUsers")); assert.ok(w.document.querySelector('input[data-k="cUsers"][list]')); assert.strictEqual(w.document.querySelector('select[data-k="cClientName"]').options.length, 3); w.$("colsCore").click(); assert.ok(!w.TL.visibleCols.has("teams") && w.TL.visibleCols.has("priority")); });
    await t("choosing a project refreshes the ticket-type list and shows the default", async () => { const tr = w.document.querySelector("#rows tr"); setVal(w, tr.querySelector('select[data-k="projectId"]'), "p1"); const g = tr.querySelector('select[data-k="groupId"]'); eq([...g.options].map((o) => o.textContent), ["Default: Task", "Task", "Bug"]); eq([...tr.querySelector('select[data-k="columnId"]').options].map((o) => o.textContent), ["Default: To Do", "To Do", "In Progress", "Done"]); });
    await t("sample sheet button decodes the embedded xlsx and offers it to save", async () => { w.window.__saved = null; w.document.querySelector("#edit .dlSample").click(); await settle(w, 40); assert.ok(w.window.__saved); assert.ok(w.window.__saved.data.size > 10000, "embedded workbook decoded"); assert.match(w.window.__saved.filename, /sample\)\.xlsx$/); });
    await t("no JS errors", async () => assert.deepStrictEqual(w.errors, []));
  }

  console.log("Inside the claude.ai sandbox");
  {
    const w = boot({ sandbox: true });
    await t("banner explains the sandbox and offers to save the app", async () => { assert.ok(w.TL.IN_SANDBOX); assert.ok(!w.$("sandboxNote").hidden); assert.match(w.$("loginBtn").textContent, /on the website/); assert.ok(w.$("sandboxNote").querySelector("a[href*='github.io/tms-ticket-loader']")); });
    await t("a sign-in attempt explains the real cause instead of blaming credentials", async () => { w.window.fetch = async () => { throw new TypeError("blocked by CSP"); }; w.$("user").value = "akash"; w.$("pass").value = "pw"; w.$("loginForm").dispatchEvent(new w.window.Event("submit", { cancelable: true })); await settle(w, 30); assert.match(w.$("loginErr").textContent, /cannot reach TMS from inside claude\.ai/); });
    await t("Save the app hands the full page to the downloads capability as .html", async () => { w.$("saveApp").click(); await settle(w, 40); assert.ok(w.window.__saved); assert.strictEqual(w.window.__saved.filename, "NUPI Ticket Loader.html"); assert.ok(String(w.window.__saved.data).startsWith("<!doctype html>")); assert.ok(String(w.window.__saved.data).includes("acquireVsCodeApi") === false && String(w.window.__saved.data).includes("NUPI Ticket Loader")); });
    await t("no JS errors", async () => assert.deepStrictEqual(w.errors, []));
  }

  console.log("Standalone file (opened from disk)");
  {
    const w = boot();
    await t("no banner, normal sign-in button", async () => { assert.ok(!w.TL.IN_SANDBOX); assert.ok(w.$("sandboxNote").hidden); assert.strictEqual(w.$("loginBtn").textContent, "Sign in"); });
    await t("template download uses a blob link outside the sandbox", async () => { await login(w); w.$("dlTemplate").click(); await settle(w, 80); assert.ok(w.window.__saved); assert.strictEqual(w.window.__saved.filename, "NUPI Ticket Import Template.xlsx"); });
    await t("no JS errors", async () => assert.deepStrictEqual(w.errors, []));
  }

  console.log("Fallbacks when lists are restricted");
  {
    const w = boot({ noUsers: true, noProjects: true });
    await login(w);
    await t("users and projects harvested from recent tickets", async () => { eq(w.TL.ref.users.map((u) => u.name), ["Akash Tripathi", "Riya Mehta", "Vipul Sharma"]); eq(w.TL.ref.projects.map((p) => p.name), ["Backoffice", "NUPI App"]); });
    await t("no JS errors", async () => assert.deepStrictEqual(w.errors, []));
  }

  console.log("Matching, parsing and import");
  {
    const w = boot(); await login(w); const TL = w.TL;
    await t("matchName: exact, username, partial, ambiguous", async () => {
      assert.strictEqual(TL.matchName("Vipul Sharma", TL.ref.users).item.id, "u2");
      assert.strictEqual(TL.matchName("vipul", TL.ref.users).item.id, "u2");
      assert.strictEqual(TL.matchName("riya m", TL.ref.users).item.id, "u3");
      assert.strictEqual(TL.matchName("u1", TL.ref.users).item.id, "u1");
      assert.strictEqual(TL.matchName("nobody", TL.ref.users).item, null);
    });
    await t("parseDelimited handles TSV, CSV with quotes, CRLF", async () => {
      eq(TL.parseDelimited("a\tb\r\nc\td\n"), [["a", "b"], ["c", "d"]]);
      eq(TL.parseDelimited('name,desc\n"Fix, it","line1\nline2"\n'), [["name", "desc"], ["Fix, it", "line1\nline2"]]);
    });
    await t("headerToKey recognises synonyms and TMS labels", async () => { assert.strictEqual(TL.headerToKey("Assigned To"), "assignedUserId"); assert.strictEqual(TL.headerToKey("Ticket title *"), "name"); assert.strictEqual(TL.headerToKey("Due Date"), "cDueDate"); assert.strictEqual(TL.headerToKey("Date End"), "dateEnd"); assert.strictEqual(TL.headerToKey("Random"), null); });
    await t("mapImported with header row resolves people, projects, status and enum", async () => {
      const r = TL.mapImported([["Ticket", "Project", "Assign to", "Status", "Priority", "Description"], ["Add OTP cooldown", "nupi app", "Vipul", "in progress", "high", "Countdown after OTP"], ["Broken export", "Backoffice", "Unknown Person", "Backlog", "", ""]]);
      assert.strictEqual(r.rows.length, 2);
      assert.strictEqual(r.rows[0].projectId, "p1"); assert.strictEqual(r.rows[0].assignedUserId, "u2"); assert.strictEqual(r.rows[0].columnId, "c2"); assert.strictEqual(r.rows[0].priority, "High");
      assert.strictEqual(r.rows[1].assignedUserId, ""); assert.strictEqual(r.rows[1]._hint.assignedUserId.cls, "bad"); assert.strictEqual(r.rows[1].columnId, "c9");
      assert.ok(TL.visibleCols.has("priority"));
    });
    await t("mapImported without header uses positional columns", async () => { const r = TL.mapImported([["Just a title", "Backoffice", "Riya"]]); assert.strictEqual(r.rows[0].name, "Just a title"); assert.strictEqual(r.rows[0].projectId, "p2"); assert.strictEqual(r.rows[0].assignedUserId, "u3"); });
    await t("pasting TSV onto the grid adds rows and flags unresolved names", async () => {
      const dt = { getData: () => "Ticket\tProject\tAssign to\nPaste one\tNUPI App\tAkash\nPaste two\tNUPI App\tGhost\n" };
      const ev = new w.window.Event("paste", { bubbles: true, cancelable: true }); ev.clipboardData = dt; w.document.body.dispatchEvent(ev);
      assert.ok(ev.defaultPrevented); assert.strictEqual(TL.rows.length, 2); assert.strictEqual(TL.rows[0].assignedUserId, "u1");
      assert.ok(w.document.querySelector(".col-user .hint.bad").textContent.includes("Ghost"));
      assert.ok(![...w.document.querySelectorAll(".hint.bad")].some((h) => /Priority is required/.test(h.textContent)), "a required field with a default is not an error");
    });
    await t("Excel import via ExcelJS (template-shaped sheet, title block skipped)", async () => {
      const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet("Tickets"); wb.addWorksheet("Lists");
      ws.addRow(["Some title block"]); ws.addRow(["Ticket title *", "Project *", "Ticket Type", "Assign to", "Stage", "Due date", "Start", "Teams", "Tester", "Estimated hours", "Allocation", "Description"]);
      ws.addRow(["From Excel A", "Backoffice", "Task", "Riya Mehta", "Backlog", new Date(Date.UTC(2026, 10, 5)), "23/10/2026", "Mobile, Backend", "Vipul", "4.5h", "Akash; Vipul", "desc A"]); ws.addRow(["From Excel B", "NUPI App", "Bug", "", "", "not a date", "", "Ghost team", "", "", "", ""]);
      const buf = await wb.xlsx.writeBuffer();
      const r = await TL.importWorkbook(w.window.Uint8Array.from(new Uint8Array(buf)).buffer); // page-realm ArrayBuffer, as file.arrayBuffer() gives
      assert.strictEqual(r.rows.length, 2); assert.strictEqual(r.rows[0].projectId, "p2"); assert.strictEqual(r.rows[0].assignedUserId, "u3"); assert.strictEqual(r.rows[0].columnId, "c9"); assert.strictEqual(r.rows[0].cDueDate, "2026-11-05"); assert.strictEqual(r.rows[1].assignedUserId, "");
      assert.strictEqual(r.rows[0].dateStart, "2026-10-23"); eq(r.rows[0].teams, ["tm1", "tm2"]); eq(r.rows[0].cUsers, ["u2"]); assert.strictEqual(r.rows[0].cEstimateHours, 4.5); eq(r.rows[0].cResourceAllocation, ["Akash", "Vipul"]); assert.strictEqual(r.rows[0].groupId, "g3"); assert.strictEqual(r.rows[1].groupId, "g2"); assert.strictEqual(r.rows[0].priority, "Normal", "default applied to imported rows");
      assert.match(r.rows[1]._hint.cDueDate.text, /Could not read the date/); assert.match(r.rows[1]._hint.teams.text, /Ghost team/);
    });
    await t("template workbook has live dropdowns and example rows", async () => {
      const wb = await TL.buildTemplate(); const ws = wb.getWorksheet("Tickets"); const lists = wb.getWorksheet("Lists");
      assert.ok(ws && lists && wb.getWorksheet("Instructions"));
      assert.strictEqual(lists.getCell("A2").value, "Akash Tripathi"); assert.strictEqual(lists.getCell("B2").value, "NUPI App"); assert.strictEqual(lists.getCell("C2").value, "To Do");
      const dv = ws.dataValidations.model; const keys = Object.keys(dv); assert.ok(keys.length >= 3);
      const userDv = Object.values(dv).find((d) => d.formulae && d.formulae[0] === "Lists!$A$2:$A$4"); assert.ok(userDv, "user dropdown references Lists!A");
      assert.match(String(ws.getCell("A2").value), /^EXAMPLE/);
      const heads = ws.getRow(1).values.slice(1); eq(heads, ["Ticket title *", "Project *", "Ticket Type", "Assign to", "Stage", "Priority *", "Date Start", "Planned Start Date", "Date End", "Due Date", "Planned End Date", "Target Date", "Reminder", "Estimate Hours", "Hours", "Points", "Client Name", "Type", "Teams", "Tester", "Allocation", "Working on it today", "Prototype Link", "Reference Ticket Number", "Description"]);
      const listHeads = lists.getRow(1).values.slice(1); eq(listHeads, ["Users", "Projects", "Stages", "Ticket types", "Teams", "Priority", "Client Name", "Type", "Allocation", "Dates"]);
      eq(lists.getColumn(4).values.slice(2), ["Task", "Bug"], "ticket types de-duplicated across projects");
      const DL = "J"; assert.ok(Object.prototype.toString.call(lists.getCell(DL + "2").value) === "[object Date]"); assert.strictEqual(lists.getCell(DL + "2").numFmt, "yyyy-mm-dd"); assert.ok(lists.getCell(DL + "300").value, "a year of dates");
      const dueCol = heads.indexOf("Due Date") + 1; const L = String.fromCharCode(64 + dueCol); assert.ok(ws.getCell(2, dueCol).value >= ws.getCell(2, dueCol - 1).value, "example due date is not before its start date");
      const dateDv = Object.entries(dv).find(([rng]) => rng.startsWith(L + "2:"))[1]; assert.strictEqual(dateDv.type, "list"); assert.strictEqual(dateDv.formulae[0], "Lists!$J$2:$J$" + (lists.rowCount)); assert.strictEqual(dateDv.showErrorMessage, false); assert.match(dateDv.prompt, /pick a date/i);
      assert.ok(ws.getCell("A1").note && JSON.stringify(ws.getCell("A1").note).includes("Required"));
      assert.ok(JSON.stringify(ws.getCell("F1").note).includes("blank = Normal"), "priority note explains the default");
      const typeDv = Object.entries(dv).find(([rng]) => rng.startsWith("C2:"))[1]; assert.strictEqual(typeDv.formulae[0], "Lists!$D$2:$D$3");
      const testerDv = Object.entries(dv).find(([rng]) => rng.startsWith("T2:"))[1]; assert.strictEqual(testerDv.formulae[0], "Lists!$A$2:$A$4");
      const teamsDv = Object.entries(dv).find(([rng]) => rng.startsWith("S2:"))[1]; assert.strictEqual(teamsDv.formulae[0], "Lists!$E$2:$E$3");
      assert.strictEqual(ws.getCell("F2").value, "Normal", "example row carries the default priority"); assert.strictEqual(ws.getCell("C2").value, "Task");
      assert.ok(wb.getWorksheet("Instructions").getCell("A9").value.toString().includes("Column guide"));
      const buf = await wb.xlsx.writeBuffer(); assert.ok(buf.byteLength > 5000);
    });
    await t("Download template calls the downloads capability with an xlsx", async () => { w.$("dlTemplate").click(); await settle(w, 80); assert.ok(w.window.__saved); assert.strictEqual(w.window.__saved.filename, "NUPI Ticket Import Template.xlsx"); });
    await t("no JS errors", async () => assert.deepStrictEqual(w.errors, []));
  }

  console.log("Validation, review and creation");
  {
    const w = boot(); await login(w); const TL = w.TL;
    await t("review with an empty row is blocked with a message", async () => { w.$("reviewBtn").click(); assert.match(w.$("editErr").textContent, /fixing/); assert.ok(w.$("review").hidden); assert.ok(w.document.querySelector("#rows tr.bad")); });
    await t("validate: required fields, duplicates in batch, wrong-board status", async () => {
      TL.rows = [TL.newRow({ name: "Ticket A long", projectId: "p1", columnId: "c9" }), TL.newRow({ name: "ticket a LONG", projectId: "p1" }), TL.newRow({ name: "ok", projectId: "" })];
      const bad = TL.validate(TL.rows); eq(bad, [0, 1, 2]);
      assert.ok(TL.rows[0]._errs.some((e) => /stage does not exist/.test(e))); assert.ok(TL.rows[1]._errs.some((e) => /Same title as row 1/.test(e))); assert.ok(TL.rows[2]._errs.some((e) => /Project is required/.test(e)));
      TL.rows = [TL.newRow({ name: "Wrong type here", projectId: "p2", groupId: "g2" })]; TL.validate(TL.rows); assert.ok(TL.rows[0]._errs.some((e) => /ticket type does not exist/.test(e)));
    });
    await t("payload carries names alongside ids and extra fields", async () => {
      const p = TL.payloadFor(TL.newRow({ name: " Add OTP cooldown ", projectId: "p1", groupId: "g2", assignedUserId: "u2", columnId: "c2", description: "d", priority: "High", cEstimateHours: "4.5", cUsers: ["u3"], teams: ["tm1"], cReminderAt: "2026-10-23T09:00", cResourceAllocation: ["Akash"] }));
      const utc = new w.window.Date("2026-10-23T09:00").toISOString().slice(0, 19).replace("T", " ");
      const sorted = (o) => Object.fromEntries(Object.entries(o).sort(([a], [b]) => a.localeCompare(b)));
      eq(sorted(p), sorted({ name: "Add OTP cooldown", projectId: "p1", projectName: "NUPI App", ownerId: "u1", ownerName: "Akash Tripathi", groupId: "g2", groupName: "Bug", assignedUserId: "u2", assignedUserName: "Vipul Sharma", columnId: "c2", columnName: "In Progress", boardId: "b1", description: "d", priority: "High", type: "Task", cUsersIds: ["u3"], cUsersNames: { u3: "Riya Mehta" }, teamsIds: ["tm1"], teamsNames: { tm1: "Mobile" }, cReminderAt: utc, cEstimateHours: 4.5, cResourceAllocation: ["Akash"] }));
      const q = TL.payloadFor(TL.newRow({ name: "Bare minimum ticket", projectId: "p2" }));
      eq(sorted(q), sorted({ name: "Bare minimum ticket", projectId: "p2", projectName: "Backoffice", ownerId: "u1", ownerName: "Akash Tripathi", groupId: "g3", groupName: "Task", columnId: "c9", columnName: "Backlog", boardId: "b2", priority: "Normal", type: "Task" }), "defaults: creator, Task type, first stage, board, Normal priority");
    });
    await t("review shows tiles, plan and groups per project", async () => {
      TL.rows = [TL.newRow({ name: "Add OTP cooldown", projectId: "p1", assignedUserId: "u2", columnId: "c2", description: "Countdown" }), TL.newRow({ name: "Fix export timeout", projectId: "p2", description: "Background job" }), TL.newRow({ name: "Existing ticket one", projectId: "p1", assignedUserId: "u1" })];
      w.$("reviewBtn").click(); assert.ok(!w.$("review").hidden); assert.ok(w.$("edit").hidden);
      assert.match(w.$("plan").textContent, /default type \(Task\) and you \(Akash Tripathi\) as Creator/); assert.match(w.$("plan").textContent, /Priority defaults to Normal/);
      assert.ok([...w.$("groups").querySelectorAll("tbody td:nth-child(2)")].every((td) => td.textContent === "Task"), "type column shows the default");
      assert.match(w.$("tiles").textContent, /Tickets3/); assert.match(w.$("plan").textContent, /Create 3 ProjectTask/); assert.match(w.$("plan").textContent, /1 ticket will be created unassigned/);
      assert.strictEqual(w.$("groups").querySelectorAll(".group").length, 2); assert.strictEqual(w.$("confirmCount").textContent, "3");
    });
    await t("duplicate check against TMS marks the existing title", async () => { w.$("dupCheck").click(); await settle(w, 40); assert.match(w.$("dupStatus").textContent, /1 ticket already exists/); assert.ok(w.$("groups").querySelector(".pill.warn").textContent.includes("already in TMS")); });
    await t("confirm creates each ticket with the right payload and links them", async () => {
      w.$("confirmBtn").click(); await settle(w, 80);
      assert.strictEqual(w.fake.created.length, 3); assert.strictEqual(w.fake.created[0].assignedUserName, "Vipul Sharma"); assert.strictEqual(w.fake.created[0].columnName, "In Progress"); assert.strictEqual(w.fake.created[1].columnName, "Backlog", "blank stage -> project's first stage"); assert.strictEqual(w.fake.created[1].groupName, "Task"); assert.strictEqual(w.fake.created[1].ownerId, "u1"); assert.strictEqual(w.fake.created[1].priority, "Normal");
      assert.match(w.$("resTitle").textContent, /3 tickets created/); assert.strictEqual(w.$("resRows").querySelectorAll("a.pill.ok").length, 3); assert.ok(w.$("resRows").querySelector("a").href.includes("#ProjectTask/view/n1"));
      assert.ok(w.$("retryBtn").hidden); assert.ok(!w.$("resActions").hidden);
    });
    await t("results CSV has one line per ticket with links", async () => { const csv = TL.resultsCsv(); assert.strictEqual(csv.split("\r\n").length, 4); assert.ok(csv.includes("#ProjectTask/view/n2")); });
    await t("new batch resets to one empty row", async () => { w.$("newBatch").click(); assert.ok(!w.$("edit").hidden); assert.strictEqual(TL.rows.length, 1); assert.strictEqual(TL.rows[0].name, ""); });
    await t("no JS errors", async () => assert.deepStrictEqual(w.errors, []));
  }

  console.log("Failures and retry");
  {
    const w = boot({ failName: "Will fail", failOnce: 0 }); await login(w); const TL = w.TL;
    await t("a refused ticket is reported, the rest succeed, retry fixes it", async () => {
      TL.rows = [TL.newRow({ name: "Will succeed", projectId: "p1" }), TL.newRow({ name: "Will fail", projectId: "p2" })];
      w.$("reviewBtn").click(); w.$("confirmBtn").click(); await settle(w, 80);
      assert.match(w.$("resTitle").textContent, /1 created, 1 failed/); assert.ok(!w.$("retryBtn").hidden); assert.match(w.$("resRows").textContent, /cannot create tickets in this project/);
      w.$("retryBtn").click(); await settle(w, 80);
      assert.match(w.$("resTitle").textContent, /2 tickets created/); assert.strictEqual(w.fake.created.length, 2);
    });
    await t("no JS errors", async () => assert.deepStrictEqual(w.errors, []));
  }

  console.log("Drafts and sign-out");
  {
    const w = boot(); await login(w); const TL = w.TL;
    await t("edits are kept as a draft and restored after reload", async () => {
      const tr = w.document.querySelector("#rows tr"); setVal(w, tr.querySelector('input[data-k="name"]'), "Draft ticket title"); setVal(w, tr.querySelector('select[data-k="projectId"]'), "p1");
      const draft = JSON.parse(w.window.localStorage.getItem("tl.draft")); assert.strictEqual(draft.rows[0].name, "Draft ticket title");
      TL.rows = []; await TL.afterLogin(); assert.strictEqual(TL.rows[0].name, "Draft ticket title"); assert.ok(!w.$("draftNote").hidden);
    });
    await t("sign out destroys the token and returns to sign-in", async () => { w.$("signOut").click(); await settle(w, 20); assert.ok(w.fake.calls.some((c) => c.path === "App/action/destroyAuthToken")); assert.ok(!w.$("signin").hidden); assert.strictEqual(w.window.sessionStorage.getItem("tl.session"), null); });
    await t("no JS errors", async () => assert.deepStrictEqual(w.errors, []));
  }

  console.log("\n" + passes + " passed, " + failures + " failed");
  if (failures) { process.exitCode = 1; } else console.log("all tests passed");
})();
