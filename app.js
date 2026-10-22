/* ATLAS OS v1.0.0 — by AashiFX
   CHIEF › Kingdom › Unit › Record. Morning Brief, Approval Inbox, Command Bar, Record Registry.
   Data: sample data until Microsoft sign-in is configured (config.js), then live from OneDrive via Microsoft Graph.
   This repository contains no business data. */
(function () {
  "use strict";
  var VERSION = "1.3.5";
  var CFG = window.ATLAS_CONFIG || {};

  // ---------------------------------------------------------------- helpers
  function $(id) { return document.getElementById(id); }
  function esc(v) { return String(v == null ? "" : v).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
  function num(v) { var n = typeof v === "number" ? v : parseFloat(String(v || "").replace(/,/g, "")); return isFinite(n) ? n : 0; }
  function sar(v, dp) { dp = dp == null ? 2 : dp; return num(v).toLocaleString("en-US", { minimumFractionDigits: dp, maximumFractionDigits: dp }); }
  var MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  var DAY = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  function toDate(v) {
    if (v == null || v === "") return null;
    if (typeof v === "number" && v > 20000 && v < 80000) return new Date(Math.round((v - 25569) * 864e5));
    var s = String(v).trim(), m;
    if ((m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})/))) { var mi = MON.indexOf(m[2].charAt(0).toUpperCase() + m[2].slice(1, 3).toLowerCase()); if (mi >= 0) return new Date(Date.UTC(+m[3], mi, +m[1])); }
    if ((m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/))) return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
    return null;
  }
  function fmt(d) { return d ? ("0" + d.getUTCDate()).slice(-2) + "-" + MON[d.getUTCMonth()] + "-" + d.getUTCFullYear() : "—"; }
  function fmtS(d) { return d ? ("0" + d.getUTCDate()).slice(-2) + "-" + MON[d.getUTCMonth()] : "—"; }
  function todayUTC() { var n = new Date(); return Date.UTC(n.getFullYear(), n.getMonth(), n.getDate()); }
  function daysTo(d) { return d ? Math.round((d.getTime() - todayUTC()) / 864e5) : null; }
  function vName(plate, key) { var p = String(plate || "").split("|").map(function (x) { return x.trim(); }).filter(Boolean); return p.length >= 2 ? p.slice(0, -1).join(" ") + " · " + p[p.length - 1] : (plate || key || ""); }
  function store(k, v) { try { if (v === undefined) return JSON.parse(localStorage.getItem("atlas." + k) || "null"); localStorage.setItem("atlas." + k, JSON.stringify(v)); } catch (e) { return null; } }
  function icon(d, s) { return '<svg aria-hidden="true" width="' + (s || 18) + '" height="' + (s || 18) + '" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="' + d + '"/></svg>'; }

  // ---------------------------------------------------------------- structure (10 approved kingdoms)
  var I = {
    home: "M3 11l9-7 9 7M5 10v10h14V10", brief: "M12 3v2M5.6 5.6l1.4 1.4M3 12h2M19 12h2M17 7l1.4-1.4M7 16a5 5 0 1 1 10 0zM4 19h16",
    inbox: "M3 13l3-8h12l3 8v6H3zM3 13h5l1 3h6l1-3h5", registry: "M5 4h11l3 3v13H5zM9 9h7M9 13h7M9 17h4",
    atlas: "M12 3l8 4v6c0 4-3.5 7-8 8-4.5-1-8-4-8-8V7z", kb: "M4 5a2 2 0 0 1 2-2h13v16H6a2 2 0 0 0-2 2zM4 19a2 2 0 0 1 2-2h13",
    arch: "M4 20h16M6 20V9l6-5 6 5v11M10 20v-6h4v6", biz: "M3 7h18v13H3zM8 7V4h8v3M3 12h18", write: "M4 20l4-1 11-11-3-3L5 16zM14 6l3 3",
    creative: "M12 3a9 9 0 1 0 0 18c1 0 2-1 1-2-1-2 0-3 2-3h2a4 4 0 0 0 4-4 9 9 0 0 0-9-9z", lab: "M9 3h6M10 3v6l-5 9a2 2 0 0 0 2 3h10a2 2 0 0 0 2-3l-5-9V3",
    personal: "M12 12a4 4 0 1 0 0-8 4 4 0 0 0 0 8M4 21v-1a6 6 0 0 1 6-6h4a6 6 0 0 1 6 6v1", dawah: "M12 3a9 9 0 1 0 9 9 7 7 0 0 1-9-9z",
    learn: "M2 9l10-5 10 5-10 5zM6 11v5c3 2 9 2 12 0v-5"
  };
  // Record IDs follow CHIEF's Evidence ID format [Company]-[Dept]-[Seq]. Proposed — confirm in Registry.
  function R(id, kind, name, about, status, source) { return { id: id, kind: kind, name: name, about: about, status: status, source: source || null }; }
  var KINGDOMS = [
    { id: "atlas", name: "ATLAS™", icon: I.atlas, about: "Executive intelligence & coordination — routing, handovers, daily command.", units: [
      { id: "desk", name: "Command Desk", about: "Daily priorities, handovers, routing", records: [R("AFX-EXE-001", "LOG", "Handover register", "“ATLAS, handover” records", "PLANNED")] }] },
    { id: "kb", name: "AASHIF OS™ Knowledge Base", icon: I.kb, about: "Institutional repository — authoritative records and registration rules.", units: [
      { id: "repo", name: "Repository", about: "Seven volumes · Evidence IDs · master forms", records: [R("AFX-DOC-001", "DOC", "AASHIF OS™ Volumes I–VII", "Constitution → Reference Library", "TO CONNECT")] }] },
    { id: "arch", name: "CHIEF Architect™", icon: I.arch, about: "Architecture, governance and structural decisions.", units: [
      { id: "office", name: "Architecture Office", about: "Blueprints and decisions", records: [R("AFX-EXE-002", "DOC", "ATLAS OS Vision Blueprint", "Kingdoms, roadmap, decisions", "TO CONNECT")] }] },
    { id: "biz", name: "Business Operations", icon: I.biz, about: "Jama Star · MAS Mobility · Prestige Rental — operations, billing, accounts.", unitLabel: "Business units", units: [
      { id: "jama", name: "Jama Star Rent a Car", about: "Monthly & daily rentals · KAEC and WWC", records: [
        R("JAM-FIN-001", "XLS", "Jama Client Accounts", "Invoices, payments, SOA · KAEC + WWC", "TO CONNECT", "jamaClientAccounts"),
        R("JAM-FIN-002", "XLS", "Jama Star Expense & VAT Register", "Expenses · VAT REGISTER · quarterly VAT return", "TO CONNECT", "vatHub"),
        R("JAM-FIN-003", "XLS", "ADIF payment allocation", "Monthly allocation records", "TO CONNECT"),
        R("JAM-FLT-001", "XLS", "Vehicle Maintenance Record", "Fleet service history", "TO CONNECT"),
        R("JAM-FIN-004", "XLS", "Monthly Finance Report", "Management accounts · monthly DUE / EXCESS", "TO CONNECT", "monthlyFinance")] },
      { id: "mas", name: "MAS Mobility", about: "Premium chauffeur & charter transport", records: [R("MAS-FIN-001", "XLS", "MAS Accounts", "Main Statement · documents · receipts", "TO CONNECT", "masAccounts")] },
      { id: "prestige", name: "Prestige Rental Equipment", about: "Heavy equipment rental · quotations", records: [
        R("PRE-SAL-001", "XLS", "Quotation Preparation", "RFQs · PR26-Q09 standard", "TO CONNECT"),
        R("PRE-PRO-001", "DOC", "SELA project", "Project documents", "TO CONNECT")] }] },
    { id: "write", name: "Writing & Communication Studio", icon: I.write, about: "Letters, emails, reports and official communication.", units: [
      { id: "studio", name: "Studio Desk", about: "Drafts and templates", records: [R("AFX-COM-001", "DOC", "Letterheads & templates", "Bilingual document standards", "TO CONNECT")] }] },
    { id: "creative", name: "Creative Studio", icon: I.creative, about: "Posters, branding, logos and final visuals.", units: [
      { id: "vault", name: "Brand Vault", about: "ATLAS™ seal, AashiFX assets", records: [R("AFX-DOC-002", "IMG", "ATLAS Ultra Steel Masters", "Approved seal files", "TO CONNECT")] }] },
    { id: "lab", name: "AI & Automation Lab", icon: I.lab, about: "Office Scripts, Power Automate, Python, ATLAS OS build.", units: [
      { id: "library", name: "Automation Library", about: "Scripts and flows", records: [R("AFX-IT-001", "TXT", "Jama Office Scripts", "Accounts Actions, SOA, Daily Services", "TO CONNECT"), R("AFX-IT-002", "APP", "ATLAS OS", "This application · v" + VERSION, "LIVE")] }] },
    { id: "personal", name: "Personal Executive", icon: I.personal, about: "Personal administration, schedules and household matters.", units: [
      { id: "pdesk", name: "Personal Desk", about: "Personal records", records: [R("AFX-ADM-001", "DOC", "Personal records", "Point ATLAS to your files", "TO CONNECT")] }] },
    { id: "dawah", name: "Islamic Growth & Dawah", icon: I.dawah, about: "Daily Qur'an & Hadith bulletin and dawah work.", units: [
      { id: "bulletin", name: "Daily Bulletin", about: "Approved multilingual structure", records: [R("AFX-DOC-003", "DOC", "Qur'an & Hadith bulletin", "Daily issue", "TO CONNECT")] }] },
    { id: "learn", name: "Learning & Career Development", icon: I.learn, about: "Certifications, study plans and career records.", units: [
      { id: "track", name: "Learning Track", about: "Courses and certificates", records: [R("AFX-DOC-004", "DOC", "Google AI Professional Certificate", "Enrolment & progress", "TO CONNECT")] }] }
  ];
  var RECORDS = [];
  KINGDOMS.forEach(function (k) { k.units.forEach(function (u) { u.records.forEach(function (r) { r.k = k; r.u = u; RECORDS.push(r); }); }); });
  function findK(id) { return KINGDOMS.filter(function (k) { return k.id === id; })[0]; }
  function findU(k, id) { return k && k.units.filter(function (u) { return u.id === id; })[0]; }
  function findR(id) { return RECORDS.filter(function (r) { return r.id === id; })[0]; }

  // ---------------------------------------------------------------- state
  var S = { mode: null, hubs: {}, errors: {}, files: {}, theme: store("theme") || "dark", data: null, live: false, error: null, account: null, done: store("done") || {}, log: store("log") || [] };
  function logAct(text) { S.log.unshift({ t: new Date().toISOString(), text: text }); S.log = S.log.slice(0, 200); store("log", S.log); }

  // ---------------------------------------------------------------- data model
  // Source of truth = workbook TABLES (not the dashboard views, which change with the selected vehicle):
  //   tblAccountRegister → one row per client account (unit, source record, customer, vehicle, plate, start, end, rate, status)
  //   KAEC / WWC vehicle-record tables → billing cycles per source record
  //   Unified Control Center rows → WWC daily-service balances
  function isActive(v) { return /^active$/i.test(String(v.status || "").trim()); }
  function model(raw) {
    function recs(rows, unit) {
      return (rows || []).filter(function (r) { return r[0] !== "" && r[0] != null; }).map(function (r) {
        return { unit: unit, key: String(r[0]).trim(), no: r[1], start: toDate(r[2]), end: toDate(r[3]), period: typeof r[3] === "string" ? r[3] : "", inv: r[4] || "", invStatus: String(r[5] || ""), debit: num(r[6]), credit: num(r[7]), bal: num(r[8]), payStatus: String(r[11] || "") };
      });
    }
    var kRec = recs(raw.kRec, "KAEC"), wRec = recs(raw.wRec, "WWC");
    var issued = function (c) { return /issued/i.test(c.invStatus) && !/not/i.test(c.invStatus); };
    var accounts = [];
    (raw.ar || []).forEach(function (r) {
      var unit = String(r[0] || "").trim(), key = String(r[1] || "").trim(), cust = String(r[8] || "").trim();
      if (!unit || !key || !cust || /no current/i.test(cust)) return;
      var start = toDate(r[11]), end = toDate(r[12]), status = String(r[15] || "").trim();
      var src = unit === "WWC" ? wRec : kRec;
      var cyc = src.filter(function (c) { return c.key === key && (!start || !c.start || c.start.getTime() >= start.getTime() - 864e5) && (!end || !c.start || c.start.getTime() < end.getTime()); });
      var iss = cyc.filter(issued), open = cyc.filter(function (c) { return !issued(c); });
      var out = iss.reduce(function (s, c) { return s + c.bal; }, 0), oldest = null;
      iss.forEach(function (c) { if (c.bal > 0 && c.start && (!oldest || c.start < oldest)) oldest = c.start; });
      var veh = String(r[9] || "").trim(), plate = String(r[10] || "").trim();
      accounts.push({ id: key + "@" + (start ? start.getTime() / 864e5 : 0), unit: unit, key: key, customer: cust, name: vName(veh && plate ? veh + " | " + plate : veh || plate, key),
        start: start, end: end, rate: num(r[13]), status: status || "—", cycles: cyc, out: out, oldest: oldest,
        next: isActive({ status: status }) && open.length ? open[0].start : null, billed: iss.reduce(function (s, c) { return s + c.debit; }, 0), received: iss.reduce(function (s, c) { return s + c.credit; }, 0), issuedCount: iss.length });
    });
    (raw.att || []).forEach(function (r) {
      var key = String(r[1] || "").trim();
      if (r[0] !== "WWC" || !/^Daily /.test(key)) return;
      accounts.push({ id: key, unit: "WWC", key: key, daily: true, customer: String(r[2] || ""), name: vName(r[3], key), rate: 0, status: num(r[8]) > 0 ? "Active" : "Paid",
        cycles: [], out: num(r[8]), oldest: null, next: null, billed: 0, received: 0, issuedCount: 1 });
    });
    var kaec = accounts.filter(function (a) { return a.unit === "KAEC"; }), wwc = accounts.filter(function (a) { return a.unit === "WWC" && (isActive(a) || a.out > 0); });
    return { accounts: accounts, kaec: kaec, active: kaec.filter(isActive), wwc: wwc, at: new Date(), webUrl: raw.webUrl || null };
  }
  function sample() {
    // Fictional sample data — real figures come only from your own workbook after sign-in.
    var t = todayUTC() / 864e5 + 25569, ar = [], kRec = [], wRec = [];
    [["Sample A 1001", "Sample Customer A", "Sample Sedan", "1001 AAA", 2400, 2], ["Sample B 1002", "Sample Customer B", "Sample SUV", "1002 BBB", 2400, 4],
      ["Sample C 1003", "Sample Customer C", "Sample Sedan", "1003 CCC", 2100, 4], ["Sample D 1004", "Sample Customer D", "Sample Crossover", "1004 DDD", 2500, 12],
      ["Sample E 1005", "Sample Customer E", "Sample Hatchback", "1005 EEE", 2200, 27]].forEach(function (v) {
      ar.push(["KAEC", v[0], "", "", "", "", "", "", v[1], v[2], v[3], t + v[5] - 90, "", v[4], "", "Active"]);
      kRec.push([v[0], 1, t + v[5] - 90, t + v[5] - 60, "SAMPLE-1", "Issued", v[4], v[4], 0, "", "", "Paid"], [v[0], 2, t + v[5] - 60, t + v[5] - 30, "SAMPLE-2", "Issued", v[4], v[4], 0, "", "", "Paid"],
        [v[0], 3, t + v[5] - 30, t + v[5], "SAMPLE-3", "Issued", v[4], v[4], 0, "", "", "Paid"], [v[0], 4, t + v[5], t + v[5] + 30, "", "Not Issued", "", "", "", "", "", ""]);
    });
    ar.push(["WWC", "Sample W 2001", "", "", "", "", "", "", "Sample Account W", "Sample Sedan", "2001 WWW", t - 70, "", 2645, "", "Active"], ["WWC", "Sample X 2002", "", "", "", "", "", "", "Sample Account X", "Sample SUV", "2002 XXX", t - 10, "", 2645, "", "Active"]);
    wRec.push(["Sample W 2001", 1, t - 70, t - 40, "SAMPLE-W1", "Issued", 2645, 2645, 0, "", "", "Paid"], ["Sample W 2001", 2, t - 40, t - 10, "SAMPLE-W2", "Issued", 2645, 0, 2645, "", "", "Unpaid - Overdue"],
      ["Sample W 2001", 3, t - 10, t + 20, "SAMPLE-W3", "Issued", 2645, 0, 2645, "", "", "Unpaid"], ["Sample W 2001", 4, t + 20, t + 50, "", "Not Issued", "", "", "", "", "", ""],
      ["Sample X 2002", 1, t - 10, t + 20, "SAMPLE-X1", "Issued", 2645, 0, 2645, "", "", "Unpaid"], ["Sample X 2002", 2, t + 20, t + 50, "", "Not Issued", "", "", "", "", "", ""]);
    var att = [["WWC", "Daily Sample 3001", "Sample Daily Client", "Sample Sedan | 3001 DDD", 0, "", "", "", 350, ""]];
    return { ar: ar, kRec: kRec, wRec: wRec, att: att };
  }

  // ---------------------------------------------------------------- Microsoft sign-in + Graph (active once msClientId is set)
  var msal = null;
  function loadScript(src) { return new Promise(function (ok, bad) { var s = document.createElement("script"); s.src = src; s.onload = ok; s.onerror = function () { bad(new Error("Could not load " + src)); }; document.head.appendChild(s); }); }
  function msReady() {
    if (!CFG.msClientId) return Promise.resolve(null);
    if (msal) return Promise.resolve(msal);
    return loadScript("https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.38.3/lib/msal-browser.min.js").then(function () {
      msal = new window.msal.PublicClientApplication({ auth: { clientId: CFG.msClientId, authority: "https://login.microsoftonline.com/" + (CFG.msTenant || "organizations"), redirectUri: location.origin + "/" }, cache: { cacheLocation: "localStorage" } });
      return msal.handleRedirectPromise().then(function (r) { if (r && r.account) msal.setActiveAccount(r.account); var a = msal.getActiveAccount() || msal.getAllAccounts()[0]; if (a) { msal.setActiveAccount(a); S.account = a; } return msal; });
    });
  }
  function token() {
    var a = msal.getActiveAccount();
    return msal.acquireTokenSilent({ scopes: CFG.scopes, account: a }).then(function (r) { return r.accessToken; })
      .catch(function () { return msal.acquireTokenRedirect({ scopes: CFG.scopes }); });
  }
  function graph(path, tk) {
    return fetch("https://graph.microsoft.com/v1.0" + path, { headers: { Authorization: "Bearer " + tk } }).then(function (r) {
      if (!r.ok) return r.text().then(function (t) { throw new Error("Microsoft Graph " + r.status + ": " + t.slice(0, 160)); });
      return r.json();
    });
  }
  function spec(key) { var s = (CFG.records || {})[key]; return typeof s === "string" ? { names: [s] } : (s || { names: [] }); }
  function findFile(key, tk) {
    var sp = spec(key), sel = "?$select=id,name,webUrl,lastModifiedDateTime,parentReference";
    var byPath = sp.path ? graph("/me/drive/root:/" + sp.path.split("/").map(encodeURIComponent).join("/") + sel, tk).catch(function () { return null; }) : Promise.resolve(null);
    return byPath.then(function (f) {
      if (f && f.id) return f;
      var names = sp.names || [], i = 0;
      function next() {
        if (i >= names.length) return null;
        var name = names[i++];
        return graph("/me/drive/root/search(q='" + encodeURIComponent(name.replace(/\.xlsx$/i, "").replace(/'/g, "''")) + "')" + sel, tk).then(function (res) {
          // several copies can share a name — use the most recently saved exact match
          var hit = (res.value || []).filter(function (x) { return x.name === name; }).sort(function (a, b) { return String(b.lastModifiedDateTime).localeCompare(String(a.lastModifiedDateTime)); })[0];
          return hit || next();
        });
      }
      return next();
    });
  }
  function readHub(key, tk) {
    return findFile(key, tk).then(function (f) {
      if (!f) throw new Error("file not found in OneDrive (" + (spec(key).names || []).join(" / ") + ")");
      S.files[key] = f;
      return graph("/me/drive/items/" + f.id + "/workbook/worksheets/" + encodeURIComponent(CFG.hubSheet || "Hub Status") + "/range(address='" + (CFG.hubRange || "A4:C40") + "')?$select=values", tk)
        .catch(function () { throw new Error("sheet “" + (CFG.hubSheet || "Hub Status") + "” not found in " + f.name); })
        .then(function (x) {
          var keys = {};
          (x.values || []).slice(1).forEach(function (r) { var k = String(r[0] || "").trim(); if (k) keys[k] = { v: r[1], note: String(r[2] || "") }; });
          return { status: "live", file: f, keys: keys, readAt: new Date() };
        });
    });
  }
  function readJamaLive(tk) {
    return findFile("jamaClientAccounts", tk).then(function (f) {
        if (!f) throw new Error("“Jama Client Accounts.xlsx” was not found in your OneDrive.");
        S.files.jamaClientAccounts = f;
        var wb = "/me/drive/items/" + f.id + "/workbook/";
        return graph(wb + "tables?$select=name", tk).then(function (t) {
          var names = (t.value || []).map(function (x) { return x.name; });
          var pick = function (re) { return names.filter(function (n) { return re.test(n); })[0]; };
          var body = function (n) { return n ? graph(wb + "tables/" + encodeURIComponent(n) + "/dataBodyRange?$select=values", tk).then(function (x) { return x.values; }) : Promise.resolve([]); };
          if (!pick(/^tblAccountRegister$/)) throw new Error("table tblAccountRegister not found in " + f.name);
          return Promise.all([body(pick(/^tblAccountRegister$/)), body(pick(/KAECVehicleRecords/)), body(pick(/WWCVehicleRecords/)),
            graph(wb + "worksheets/" + encodeURIComponent("Unified Control Center") + "/range(address='A12:J40')?$select=values", tk).then(function (x) { return x.values; }).catch(function () { return []; })])
            .then(function (v) { return { ar: v[0], kRec: v[1], wRec: v[2], att: v[3], webUrl: f.webUrl, file: f }; });
        });
    });
  }
  function signIn() {
    if (!CFG.msClientId) { toast("Microsoft sign-in is not configured in this build."); return; }
    msReady().then(function (m) { m.loginRedirect({ scopes: CFG.scopes }); });
  }

  function setStatus(src, st) { RECORDS.forEach(function (r) { if (r.source === src) r.status = st; }); }
  function sampleHubs() {
    var now = new Date();
    return {
      masAccounts: { status: "live", sample: true, readAt: now, file: { name: "MAS Accounts (sample)" }, keys: { main_statement_balance: { v: 12345 }, main_statement_open_docs: { v: 3 }, health_status: { v: "PASS" }, last_action: { v: "Sample action" }, currency: { v: "SAR" } } },
      vatHub: { status: "live", sample: true, readAt: now, file: { name: "Expense & VAT (sample)" }, keys: { return_quarter: { v: "2026-Q3" }, return_status: { v: "OPEN" }, return_result: { v: "PAYABLE 1,000.00" }, return_due_date: { v: 46326 }, return_lines_not_selected: { v: 0 }, health_status: { v: "PASS" }, last_action: { v: "Sample" }, currency: { v: "SAR" } } },
      monthlyFinance: { status: "live", sample: true, readAt: now, file: { name: "Monthly Finance Report (sample)" }, keys: { current_month: { v: "2026-10" }, month_result: { v: "DUE 1,000.00" }, basis: { v: "BANK ONLY - sample" }, health_status: { v: "WARNINGS" }, last_action: { v: "Sample action" }, currency: { v: "SAR" } } }
    };
  }
  function load() {
    S.error = null; S.errors = {};
    if (!CFG.msClientId) { // developer preview only — never used once sign-in is configured
      S.mode = "sample"; S.live = false; S.data = model(sample()); S.hubs = sampleHubs(); setStatus("jamaClientAccounts", "LIVE"); setStatus("masAccounts", "LIVE"); setStatus("monthlyFinance", "LIVE"); setStatus("vatHub", "LIVE");
      render(); return Promise.resolve();
    }
    return msReady().then(function () {
      if (!S.account) { S.mode = "signin"; S.live = false; S.data = null; return; }
      return token().then(function (tk) {
        S.mode = "live"; S.live = true;
        var jobs = [readJamaLive(tk).then(function (raw) { S.data = model(raw); setStatus("jamaClientAccounts", "LIVE"); })
          .catch(function (e) { S.errors.jamaClientAccounts = e.message || String(e); S.data = model({}); setStatus("jamaClientAccounts", "NOT CONNECTED"); })];
        ["masAccounts", "monthlyFinance", "vatHub"].forEach(function (k) {
          jobs.push(readHub(k, tk).then(function (h) { S.hubs[k] = h; setStatus(k, "LIVE"); })
            .catch(function (e) { S.hubs[k] = { status: "not connected", reason: e.message || String(e) }; setStatus(k, "NOT CONNECTED"); }));
        });
        jobs.push(readHub("jamaClientAccounts", tk).then(function (h) { S.hubs.jamaClientAccounts = h; }).catch(function () { S.hubs.jamaClientAccounts = null; }));
        return Promise.all(jobs);
      });
    }).catch(function (e) { S.mode = "error"; S.live = false; S.error = e.message || String(e); }).then(render);
  }

  // ---------------------------------------------------------------- derived: Morning Brief + Approval Inbox
  function briefItems() {
    var d = S.data, out = [];
    d.active.slice().sort(function (a, b) { return a.next - b.next; }).forEach(function (v) {
      var n = daysTo(v.next); if (n == null || n > 7) return;
      out.push({ k: "biz", tone: n < 0 ? "bad" : "warn", tag: n < 0 ? "OVERDUE " + (-n) + "D" : n === 0 ? "DUE TODAY" : "DUE IN " + n + "D", title: "Invoice · " + v.name, sub: v.customer + " · SAR " + sar(v.rate, 0) + " · " + fmtS(v.next), go: "#/inbox" });
    });
    var wOut = d.wwc.reduce(function (s, a) { return s + a.out; }, 0);
    if (wOut > 0) out.push({ k: "biz", tone: "bad", tag: "COLLECTIONS", title: "WWC · SAR " + sar(wOut) + " outstanding", sub: d.wwc.length + " accounts — oldest first in the Inbox", go: "#/inbox" });
    var now = new Date(), q = Math.floor(now.getMonth() / 3), due = new Date(Date.UTC(now.getFullYear(), q * 3 + 1, 0));
    var qn = q === 0 ? 4 : q, qy = q === 0 ? now.getFullYear() - 1 : now.getFullYear();
    out.push({ k: "biz", tone: daysTo(due) <= 14 ? "warn" : "", tag: "VAT", title: "VAT return Q" + qn + " " + qy + " · Jama", sub: "Filing deadline " + fmt(due) + " (" + daysTo(due) + " days) — confirm with your VAT workbook", go: "#/record/JAM-FIN-002" });
    out.push({ k: "dawah", tone: "", tag: "TODAY", title: "Daily Qur'an & Hadith bulletin", sub: "Islamic Growth & Dawah", go: "#/k/dawah" });
    return out;
  }
  function inboxItems() {
    var d = S.data, out = [];
    d.active.forEach(function (v) {
      var n = daysTo(v.next); if (n == null || n > 7) return;
      out.push({ id: "inv|" + v.key + "|" + fmt(v.next), tone: n < 0 ? "bad" : "warn", kind: "INV", title: "Issue invoice · " + v.name, sub: v.customer + " · SAR " + sar(v.rate, 0) + " · billing " + fmt(v.next), sheet: "Invoice Generator", rec: "JAM-FIN-001", sort: n });
    });
    d.wwc.filter(function (a) { return a.out > 0; }).sort(function (a, b) { return (a.oldest || 9e15) - (b.oldest || 9e15); }).forEach(function (a) {
      out.push({ id: "soa|" + a.key + "|" + sar(a.out), tone: "bad", kind: "SOA", title: "Send SOA reminder · " + a.customer, sub: a.name + " · SAR " + sar(a.out) + " open" + (a.oldest ? " · oldest " + fmt(a.oldest) : ""), sheet: "SOA Center", rec: "JAM-FIN-001", sort: 50 });
    });
    return out.sort(function (a, b) { return a.sort - b.sort; });
  }
  function openCount() { if (!S.data) return 0; return inboxItems().filter(function (x) { return !S.done[x.id]; }).length; }

  // ---------------------------------------------------------------- rendering
  function route() { var h = (location.hash || "#/home").replace(/^#\//, "").split("/"); return { v: h[0] || "home", a: h[1] ? decodeURIComponent(h[1]) : null, b: h[2] ? decodeURIComponent(h[2]) : null, b2: h[3] ? decodeURIComponent(h[3]) : null }; }
  function crumbs(list) { $("crumbs").innerHTML = list.map(function (c, i) { return (i ? "<span>›</span>" : "") + '<button data-go="' + c[1] + '">' + esc(c[0]) + "</button>"; }).join(""); }
  function renderNav(r) {
    var on = function (x) { return x ? " on" : ""; }, n = openCount();
    var h = '<button class="nav-btn' + on(r.v === "home") + '" data-go="#/home">' + icon(I.home) + '<span class="l">CHIEF Home</span></button>' +
      '<button class="nav-btn' + on(r.v === "brief") + '" data-go="#/brief">' + icon(I.brief) + '<span class="l">Morning Brief</span></button>' +
      '<button class="nav-btn' + on(r.v === "inbox") + '" data-go="#/inbox">' + icon(I.inbox) + '<span class="l">Approval Inbox</span>' + (n ? '<span class="badge">' + n + "</span>" : "") + "</button>" +
      '<button class="nav-btn' + on(r.v === "registry") + '" data-go="#/registry">' + icon(I.registry) + '<span class="l">Record Registry</span></button>' +
      '<span class="nav-label">TEN KINGDOMS</span>';
    var cur = r.v === "client" ? "biz" : r.v === "k" || r.v === "u" ? r.a : r.v === "record" && findR(r.a) ? findR(r.a).k.id : null;
    KINGDOMS.forEach(function (k) { h += '<button class="nav-btn' + on(cur === k.id) + '" data-go="#/k/' + k.id + '">' + icon(k.icon) + '<span class="l">' + esc(k.name) + "</span></button>"; });
    $("nav").innerHTML = h;
  }
  function head(eyebrow, title, lead, right) { return '<div class="head"><div><div class="eyebrow">' + esc(eyebrow) + "</div><h1>" + esc(title) + "</h1>" + (lead ? '<p class="lead">' + lead + "</p>" : "") + "</div>" + (right || "") + "</div>"; }
  function kpi(l, v, n, c) { return '<div class="kpi"><span>' + esc(l) + '</span><b class="mono ' + (c || "") + '">' + esc(v) + '</b><small class="' + (c === "bad" ? "" : c || "") + '">' + esc(n) + "</small></div>"; }
  function sourceNote() {
    if (S.mode === "sample") return '<div class="notice"><span>Developer preview — fictional sample data.</span></div>';
    if (S.errors.jamaClientAccounts) return '<div class="notice"><span><b class="bad">Could not read Jama Client Accounts.xlsx:</b> ' + esc(S.errors.jamaClientAccounts) + '</span><button class="btn sm" data-act="reload">Try again</button></div>';
    return "";
  }
  // ---------------------------------------------------------------- Hub Status cards
  function hv(h, k) { return h && h.keys && h.keys[k] ? h.keys[k].v : null; }
  var DATE_KEY = /date|_from$|^current_month$|^as_of$/i, PLAIN_KEY = /docs|count|gaps|months|entries|version|posted|void|invoice|lines/i;
  function showVal(k, v) {
    if (typeof v === "number" && DATE_KEY.test(k) && v > 20000 && v < 80000) { var d = toDate(v); return /month$/i.test(k) ? MON[d.getUTCMonth()] + " " + d.getUTCFullYear() : fmt(d); }
    if (typeof v === "number") return PLAIN_KEY.test(k) ? String(v) : sar(v);
    return String(v == null ? "" : v);
  }
  function healthOf(h) {
    var st = hv(h, "health_status");
    if (st) return String(st).toUpperCase().trim();
    var m = String(hv(h, "health") || "").toUpperCase().match(/\b(PASS|WARNINGS|ERRORS)\b/);
    return m ? m[1] : "UNKNOWN";
  }
  function healthTag(h) { var x = healthOf(h); return '<span class="tag ' + (x === "PASS" ? "good" : x === "WARNINGS" ? "warn" : x === "ERRORS" ? "bad" : "") + '">' + x + "</span>"; }
  function money(v, h) { var c = String(hv(h, "currency") || "SAR"); return typeof v === "number" ? c + " " + sar(v) : (v == null || v === "" ? "—" : String(v)); }
  function asOf(h) { var a = hv(h, "as_of"); var d = a ? new Date(a) : null; return (d && !isNaN(d) ? "as of " + d.toLocaleString([], { dateStyle: "medium", timeStyle: "short" }) + " · " : "") + "read " + h.readAt.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }); }
  function hubCard(key, title, recId, body) {
    var h = S.hubs[key];
    if (!h) return "";
    if (h.status !== "live") return '<button class="tile" data-go="#/record/' + recId + '"><span class="top-row"><b>' + esc(title) + '</b><span class="tag warn">NOT CONNECTED</span></span><p>' + esc(h.reason || "") + "</p></button>";
    return '<button class="tile" data-go="#/record/' + recId + '"><span class="top-row"><b>' + esc(title) + "</b>" + healthTag(h) + "</span>" + body(h) +
      '<p>' + esc(hv(h, "last_action") ? "Last: " + hv(h, "last_action") : "") + '</p><p class="mono">' + esc(asOf(h)) + "</p></button>";
  }
  function hubCards() {
    var cards = hubCard("masAccounts", "MAS · Main Statement", "MAS-FIN-001", function (h) {
      return '<span class="kpi" style="padding:0;background:none"><b class="mono">' + esc(money(hv(h, "main_statement_balance"), h)) + "</b><small>" + esc((hv(h, "main_statement_open_docs") == null ? "—" : showVal("main_statement_open_docs", hv(h, "main_statement_open_docs"))) + " open documents") + "</small></span>";
    }) + hubCard("monthlyFinance", "Jama · Monthly result", "JAM-FIN-004", function (h) {
      var r = String(hv(h, "month_result") || "—");
      return '<span class="kpi" style="padding:0;background:none"><span>' + esc(showVal("current_month", hv(h, "current_month")) || "") + '</span><b class="mono ' + (/^DUE/i.test(r) ? "bad" : /^EXCESS/i.test(r) ? "good" : "") + '">' + esc(r) + '</b><small class="warn">Basis: ' + esc(hv(h, "basis") || "unknown") + "</small></span>";
    });
    cards += hubCard("vatHub", "Jama · VAT return", "JAM-FIN-002", function (h) {
      var r = String(hv(h, "return_result") || "—"); var n = hv(h, "return_lines_not_selected");
      return '<span class="kpi" style="padding:0;background:none"><span>' + esc(String(hv(h, "return_quarter") || "") + " · " + String(hv(h, "return_status") || "")) + '</span><b class="mono ' + (/^PAYABLE/i.test(r) ? "bad" : /^CREDIT/i.test(r) ? "good" : "") + '">' + esc(r) + '</b><small>' + esc("Due " + (showVal("return_due_date", hv(h, "return_due_date")) || "—") + (n ? " · " + n + " line(s) not ticked" : " · all lines ticked")) + "</small></span>";
    });
    var j = S.hubs.jamaClientAccounts;
    if (j && j.status === "live") cards += hubCard("jamaClientAccounts", "Jama Client Accounts", "JAM-FIN-001", function () { return ""; });
    return cards ? "<h2>Connected workbooks</h2>" + '<div class="grid">' + cards + "</div>" : "";
  }
  function hubTable(key) {
    var h = S.hubs[key];
    if (!h) return "";
    if (h.status !== "live") return '<div class="notice"><span><b class="bad">Not connected:</b> ' + esc(h.reason || "") + '</span><button class="btn sm" data-act="reload">Try again</button></div>';
    var rows = Object.keys(h.keys).map(function (k) {
      var v = h.keys[k].v, note = h.keys[k].note;
      if (k === "payable_to_company") note = (note ? note + " · " : "") + "Basis: " + (hv(h, "basis") || "unknown");
      var shown = showVal(k, v);
      return '<tr><td class="mono">' + esc(k) + '</td><td class="num mono strong">' + esc(shown) + "</td><td>" + esc(note) + "</td></tr>";
    }).join("");
    return '<div class="kpis">' + kpi("Health", healthOf(h), String(hv(h, "health_detail") || hv(h, "health") || ""), healthOf(h) === "PASS" ? "good" : healthOf(h) === "ERRORS" ? "bad" : "warn") +
      kpi("Last action", String(hv(h, "last_action") || "—"), asOf(h)) + kpi("File", h.file.name, h.sample ? "sample" : "OneDrive") + "</div>" +
      '<section class="card full"><div class="card-h"><h2>Hub Status</h2><small>' + Object.keys(h.keys).length + ' keys</small></div><div class="scroll"><table><thead><tr><th>KEY</th><th class="num">VALUE</th><th>NOTE</th></tr></thead><tbody>' + rows + "</tbody></table></div></section>";
  }
  function statusTone(s) { return s === "LIVE" ? "good" : s === "TO CONNECT" || s === "NOT CONNECTED" ? "warn" : ""; }

  function vHome() {
    var d = S.data, hr = new Date().getHours(), act = d.active;
    var hi = hr < 12 ? "Good morning" : hr < 17 ? "Good afternoon" : "Good evening";
    var due = act.filter(function (v) { var n = daysTo(v.next); return n != null && n <= 7; }).length;
    var kOut = d.kaec.reduce(function (s, v) { return s + v.out; }, 0), wOut = d.wwc.reduce(function (s, v) { return s + v.out; }, 0);
    var h = head("CHIEF Home", hi + ", CHIEF", "Ten kingdoms under one command.", '<div class="row"><button class="btn" data-go="#/brief">Morning Brief</button><button class="btn primary" data-go="#/inbox">Approval Inbox · ' + openCount() + "</button></div>") + sourceNote();
    h += '<div class="kpis">' + kpi("Jama · bills due in 7 days", due, "KAEC contracts", due ? "warn" : "good") + kpi("Jama · KAEC outstanding", "SAR " + sar(kOut), act.length + " vehicles on contract", kOut > 0 ? "bad" : "good") +
      kpi("Jama · WWC outstanding", "SAR " + sar(wOut), d.wwc.length + " accounts", wOut > 0 ? "bad" : "good") + kpi("Records connected", RECORDS.filter(function (r) { return r.status === "LIVE"; }).length + " of " + RECORDS.length, "Record Registry") + "</div>" + hubCards();
    h += '<div class="split"><section class="card"><div class="card-h"><h2>Today</h2><button class="btn sm" data-go="#/brief">Full brief</button></div>' + briefList(briefItems().slice(0, 4)) + "</section>";
    h += '<aside class="side"><h2>Recent activity</h2>' + (S.log.length ? S.log.slice(0, 4).map(function (l) { return '<div class="tip"><b>' + esc(l.text) + "</b><span>" + esc(new Date(l.t).toLocaleString()) + "</span></div>"; }).join("") : '<div class="tip"><span>Approvals and actions you take appear here.</span></div>') + "</aside></div>";
    h += "<h2>Kingdoms</h2>" + '<div class="grid">' + KINGDOMS.map(function (k) {
      var live = RECORDS.filter(function (r) { return r.k === k && r.status === "LIVE"; }).length;
      return '<button class="tile" data-go="#/k/' + k.id + '"><span class="top-row"><span class="ico">' + icon(k.icon, 19) + '</span><span class="tag ' + (live ? "good" : "") + '">' + (live ? "LIVE" : "TO CONNECT") + "</span></span><b>" + esc(k.name) + "</b><p>" + esc(k.about) + "</p></button>";
    }).join("") + "</div>";
    return h;
  }
  function briefList(items) {
    return items.map(function (b) { return '<div class="item"><span class="kind">' + icon(findK(b.k).icon, 17) + '</span><span class="txt"><span class="tag ' + b.tone + '">' + esc(b.tag) + "</span><b>" + esc(b.title) + "</b><span>" + esc(b.sub) + '</span></span><button class="btn sm" data-go="' + b.go + '">Open</button></div>'; }).join("");
  }
  function vBrief() {
    var t = new Date();
    return head("Morning Brief", DAY[t.getDay()] + ", " + t.getDate() + " " + MON[t.getMonth()] + " " + t.getFullYear(), "Everything that needs CHIEF today, across all kingdoms.") + sourceNote() +
      '<section class="card full">' + briefList(briefItems()) + "</section>" + '<p class="foot">Kingdoms add their own items here as their records are connected.</p>';
  }
  function vInbox() {
    var items = inboxItems(), open = items.filter(function (x) { return !S.done[x.id]; }), done = items.filter(function (x) { return S.done[x.id]; });
    var row = function (x, isDone) {
      return '<div class="item' + (isDone ? " done" : "") + '"><span class="kind">' + x.kind + '</span><span class="txt"><span class="tag ' + x.tone + '">' + (x.kind === "INV" ? "INVOICE DUE" : "COLLECTION") + "</span><b>" + esc(x.title) + "</b><span>" + esc(x.sub) + "</span></span>" +
        (isDone ? '<span class="tag good">APPROVED</span><button class="btn sm" data-undo="' + esc(x.id) + '">Undo</button>' : '<button class="btn sm" data-later="' + esc(x.id) + '">Later</button><button class="btn sm primary" data-approve="' + esc(x.id) + '">Approve</button>') + "</div>";
    };
    return head("Approval Inbox", open.length ? open.length + " waiting for CHIEF" : "All clear", "ATLAS prepares the work · you decide.") + sourceNote() +
      '<section class="card full">' + (open.length ? open.map(function (x) { return row(x, false); }).join("") : '<div class="empty">Nothing waiting. 🎯</div>') + "</section>" +
      (done.length ? "<h2>Approved</h2>" + '<section class="card full">' + done.map(function (x) { return row(x, true); }).join("") + "</section>" : "") +
      '<p class="foot">Approve opens the workbook in Excel and logs your decision on this device. One-click automatic run (Office Script via Power Automate) arrives in phase 2.</p>';
  }
  function vRegistry() {
    var h = head("Record Registry", RECORDS.length + " records", "Every workbook and document, filed by kingdom with an Evidence ID. IDs are proposed — confirm or change.");
    h += '<section class="card full"><div class="scroll"><table><thead><tr><th>EVIDENCE ID</th><th>RECORD</th><th>KINGDOM › UNIT</th><th>TYPE</th><th>STATUS</th></tr></thead><tbody>';
    RECORDS.forEach(function (r) { h += '<tr><td class="strong mono">' + r.id + '</td><td><button class="btn sm" data-go="#/record/' + r.id + '">' + esc(r.name) + "</button></td><td>" + esc(r.k.name) + " › " + esc(r.u.name) + "</td><td>" + r.kind + '</td><td class="tag ' + statusTone(r.status) + '">' + r.status + "</td></tr>"; });
    return h + "</tbody></table></div></section>";
  }
  function vKingdom(k) {
    if (!k) return '<div class="empty">Kingdom not found.</div>';
    return head("Kingdom", k.name, esc(k.about)) + (k.id === "biz" ? hubCards() : "") + "<h2>" + esc(k.unitLabel || "Units") + "</h2>" + '<div class="grid">' + k.units.map(function (u) {
      var live = u.records.filter(function (r) { return r.status === "LIVE"; }).length;
      return '<button class="tile" data-go="#/u/' + k.id + "/" + u.id + '"><span class="tag ' + (live ? "good" : "warn") + '">' + (live ? "LIVE" : "TO CONNECT") + "</span><b>" + esc(u.name) + "</b><p>" + esc(u.about) + '</p><p class="good">' + u.records.length + (u.records.length === 1 ? " record" : " records") + " · " + live + " live</p></button>";
    }).join("") + "</div>";
  }
  function recList(list) {
    return list.map(function (r) {
      return '<div class="item"><span class="kind">' + r.kind + '</span><span class="txt"><b>' + esc(r.name) + '</b><span class="mono">' + r.id + " · " + esc(r.about) + '</span></span><span class="tag ' + statusTone(r.status) + '">' + r.status + '</span><button class="btn sm" data-go="#/record/' + r.id + '">' + (r.status === "LIVE" ? "Open" : "Details") + "</button></div>";
    }).join("");
  }
  function vUnit(k, u) {
    if (!u) return '<div class="empty">Unit not found.</div>';
    var h = head(k.name + " › Unit", u.name, esc(u.about));
    if (k.id === "biz" && u.id === "jama") h += sourceNote() + jamaKpis() + clientAccounts(route().b2 || "all");
    return h + '<section class="card full"><div class="card-h"><h2>Records</h2><small>Excel opens only when you work on a record</small></div>' + recList(u.records) + "</section>";
  }

  // ---------------------------------------------------------------- Jama client accounts
  function allClients() {
    return S.data.accounts.map(function (a) { a.active = isActive(a); return a; });
  }
  function inTab(c, tab) {
    var live = c.active || c.out > 0;
    return tab === "kaec" ? c.unit === "KAEC" && live : tab === "wwc" ? c.unit === "WWC" && live : tab === "due" ? c.out > 0 : tab === "past" ? !live : live;
  }
  function clientAccounts(tab) {
    var list = allClients(), tabs = [["all", "Active"], ["kaec", "KAEC"], ["wwc", "WWC"], ["due", "Outstanding"], ["past", "Past"]];
    var f = list.filter(function (c) { return inTab(c, tab); })
      .sort(function (a, b) { return (b.active - a.active) || (b.out - a.out) || ((a.next || 9e15) - (b.next || 9e15)); });
    var h = '<section class="card full"><div class="card-h"><h2>Client accounts</h2><div class="row" role="tablist">' + tabs.map(function (t) {
      var n = list.filter(function (c) { return inTab(c, t[0]); }).length;
      return '<button role="tab" aria-selected="' + (t[0] === tab) + '" class="btn sm' + (t[0] === tab ? " primary" : "") + '" data-go="#/u/biz/jama/' + t[0] + '">' + t[1] + " · " + n + "</button>";
    }).join("") + '</div></div><div class="scroll"><table><thead><tr><th>CLIENT</th><th>VEHICLE</th><th>UNIT</th><th class="num">OUTSTANDING</th><th class="num">NEXT BILLING</th><th>STATUS</th></tr></thead><tbody>';
    f.forEach(function (c) {
      h += '<tr class="click" data-go="#/client/' + c.unit + "/" + encodeURIComponent(c.id) + '"><td class="strong">' + esc(c.customer) + "</td><td>" + esc(c.name) + "</td><td>" + c.unit + (c.daily ? " · Daily" : "") + '</td><td class="num mono ' + (c.out > 0 ? "bad" : "good") + '">' + sar(c.out) + '</td><td class="num">' + fmt(c.next) + '</td><td class="' + (c.out > 0 ? "bad" : c.active ? "good" : "") + '">' + esc(c.status) + "</td></tr>";
    });
    if (!f.length) h += '<tr><td colspan="6">No clients in this view.</td></tr>';
    return h + "</tbody></table></div></section>";
  }
  function findClient(unit, id) { return allClients().filter(function (x) { return x.unit === unit && (x.id === id || x.key === id); })[0]; }
  function vClient(unit, id) {
    var c = findClient(unit, id);
    if (!c) return '<div class="empty">Client account not found. <button class="btn sm" data-go="#/u/biz/jama">Back to Jama Star</button></div>';
    var cyc = c.cycles, issued = cyc.filter(function (x) { return /issued/i.test(x.invStatus) && !/not/i.test(x.invStatus); });
    var upcoming = c.active ? cyc.filter(function (x) { return !/issued/i.test(x.invStatus) || /not/i.test(x.invStatus); }).slice(0, 2) : [];
    var span = (c.start ? fmt(c.start) : "") + (c.end ? " → " + fmt(c.end) : c.start ? " → present" : "");
    var h = head("Jama Star › " + unit + " client", c.customer, esc(c.name) + " · " + unit + (c.daily ? " daily" : "") + (span ? " · " + span : "") + ' · <span class="' + (c.active ? "good" : "") + '">' + esc(c.status) + "</span>", '<div class="row"><button class="btn" data-excel="JAM-FIN-001">Statement (SOA)</button><button class="btn primary" data-excel="JAM-FIN-001">Open in Excel</button></div>') + sourceNote();
    h += '<div class="kpis">' + kpi("Monthly rate", c.rate ? "SAR " + sar(c.rate) : "—", unit) + kpi("Total billed", "SAR " + sar(c.billed), c.issuedCount + " invoices") + kpi("Total received", "SAR " + sar(c.received), "") +
      kpi("Outstanding", "SAR " + sar(c.out), c.out > 0 ? (c.oldest ? "Oldest unpaid " + fmt(c.oldest) : "Follow up") : "Fully collected", c.out > 0 ? "bad" : "good") + kpi("Next billing", fmt(c.next), c.next ? daysTo(c.next) + " days" : c.active ? "" : "Account " + c.status.toLowerCase(), c.next ? "warn" : "") + "</div>";
    h += '<section class="card full"><div class="card-h"><h2>Account history</h2><small>' + issued.length + " issued" + (upcoming.length ? " · next " + upcoming.length : "") + '</small></div><div class="scroll"><table><thead><tr><th>#</th><th>PERIOD</th><th>INVOICE</th><th class="num">DEBIT</th><th class="num">CREDIT</th><th class="num">BALANCE</th><th>STATUS</th></tr></thead><tbody>';
    issued.concat(upcoming).forEach(function (x) {
      var st = String(x.payStatus || x.invStatus), cls = /^paid/i.test(st) ? "good" : /unpaid|overdue/i.test(st) ? "bad" : /due/i.test(st) ? "warn" : "";
      var per = x.period || (x.start ? fmt(x.start) + " → " + fmt(x.end) : "—");
      h += "<tr><td>" + esc(x.no) + "</td><td>" + esc(per) + '</td><td class="strong">' + esc(x.inv || "—") + '</td><td class="num mono">' + (x.debit ? sar(x.debit) : "—") + '</td><td class="num mono">' + (x.credit ? sar(x.credit) : "—") + '</td><td class="num mono ' + (x.bal > 0 ? "bad" : "") + '">' + (x.debit ? sar(x.bal) : "—") + '</td><td class="strong ' + cls + '">' + esc(st) + "</td></tr>";
    });
    if (!cyc.length) h += '<tr><td colspan="7">' + (c.daily ? "Daily-service invoices are listed in the Daily Invoice register." : "No billing cycles found.") + "</td></tr>";
    return h + "</tbody></table></div></section>";
  }
  function jamaKpis() {
    var d = S.data, kOut = d.kaec.reduce(function (s, v) { return s + v.out; }, 0), wOut = d.wwc.reduce(function (s, v) { return s + v.out; }, 0), run = d.active.reduce(function (s, v) { return s + v.rate; }, 0);
    return '<div class="kpis">' + kpi("KAEC on contract", d.active.length, "vehicles", "good") + kpi("KAEC monthly run-rate", "SAR " + sar(run, 0), "") + kpi("KAEC outstanding", "SAR " + sar(kOut), kOut > 0 ? "Follow up" : "Fully collected", kOut > 0 ? "bad" : "good") + kpi("WWC outstanding", "SAR " + sar(wOut), d.wwc.length + " accounts", wOut > 0 ? "bad" : "good") + "</div>";
  }
  function vRecord(r) {
    if (!r) return '<div class="empty">Record not found.</div>';
    var h = head(r.k.name + " › " + r.u.name, r.name, '<span class="mono">' + r.id + "</span> · " + esc(r.about), r.source ? '<button class="btn primary" data-excel="' + r.id + '">Open in Excel</button>' : "");
    if (r.source === "jamaClientAccounts") {
      if (S.hubs.jamaClientAccounts) h += hubTable("jamaClientAccounts");
      var d = S.data, act = d.active.slice().sort(function (a, b) { return a.next - b.next; });
      h += sourceNote() + jamaKpis() + '<section class="card full"><div class="card-h"><h2>KAEC · next billing</h2></div><div class="scroll"><table><thead><tr><th>VEHICLE</th><th>CUSTOMER</th><th class="num">SAR / MONTH</th><th class="num">NEXT BILLING</th></tr></thead><tbody>';
      act.forEach(function (v) { var n = daysTo(v.next); h += '<tr><td class="strong">' + esc(v.name) + "</td><td>" + esc(v.customer) + '</td><td class="num mono">' + sar(v.rate, 0) + '</td><td class="num strong ' + (n != null && n <= 7 ? "warn" : "") + '">' + fmtS(v.next) + "</td></tr>"; });
      h += '</tbody></table></div></section><section class="card full"><div class="card-h"><h2>WWC · outstanding</h2></div><div class="scroll"><table><thead><tr><th>CUSTOMER</th><th>OLDEST UNPAID</th><th class="num">SAR</th></tr></thead><tbody>';
      d.wwc.slice().sort(function (a, b) { return (a.oldest || 9e15) - (b.oldest || 9e15); }).forEach(function (a) { h += '<tr><td class="strong">' + esc(a.customer) + "</td><td>" + fmt(a.oldest) + '</td><td class="num bad mono">' + sar(a.out) + "</td></tr>"; });
      h += "</tbody></table></div></section>";
    } else if (r.source && r.source !== "jamaClientAccounts") {
      h += hubTable(r.source);
    } else if (r.status === "LIVE") {
      h += '<div class="tip"><b>' + esc(r.name) + "</b><span>" + esc(r.about) + "</span></div>";
    } else {
      h += '<div class="notice"><span>This record isn’t connected yet. Tell ATLAS which OneDrive file or folder it is, and it will appear live here.</span></div>';
    }
    return h;
  }

  function render() {
    document.documentElement.setAttribute("data-theme", S.theme);
    var t = new Date(); $("today").textContent = DAY[t.getDay()] + ", " + t.getDate() + " " + MON[t.getMonth()] + " " + t.getFullYear();
    $("account").textContent = S.account ? (S.account.name || "CH").split(/\s+/).map(function (x) { return x[0]; }).join("").slice(0, 2).toUpperCase() : "CH";
    var r = route(); renderNav(r);
    if (S.mode === "signin" || (!S.data && S.mode !== "error")) {
      crumbs([["CHIEF", "#/home"]]);
      $("main").innerHTML = '<div class="empty" style="display:flex;flex-direction:column;align-items:center;gap:14px"><img src="seal-128.png" alt="" style="width:96px;height:96px"><h1>ATLAS OS</h1>' +
        (S.mode === "signin" ? '<p class="lead">Sign in to open your kingdoms. Only your AashiFX Microsoft account can sign in.</p>' : '<p class="lead">Starting ATLAS…</p>') +
        (CFG.msClientId ? '<button class="btn primary" data-act="signin">Sign in with Microsoft</button>' : "") + "</div>";
      return;
    }
    if (S.mode === "error") {
      crumbs([["CHIEF", "#/home"]]);
      $("main").innerHTML = '<div class="notice"><span><b class="bad">ATLAS could not start:</b> ' + esc(S.error) + '</span><button class="btn sm" data-act="reload">Try again</button><button class="btn sm primary" data-act="signin">Sign in with Microsoft</button></div>';
      return;
    }
    var k = r.a ? findK(r.a) : null, html, c = [["CHIEF", "#/home"]];
    if (r.v === "brief") { html = vBrief(); c.push(["Morning Brief", "#/brief"]); }
    else if (r.v === "inbox") { html = vInbox(); c.push(["Approval Inbox", "#/inbox"]); }
    else if (r.v === "registry") { html = vRegistry(); c.push(["Record Registry", "#/registry"]); }
    else if (r.v === "k") { html = vKingdom(k); if (k) c.push([k.name, "#/k/" + k.id]); }
    else if (r.v === "u") { var u = findU(k, r.b); html = vUnit(k, u); if (k) c.push([k.name, "#/k/" + k.id]); if (u) c.push([u.name, "#/u/" + k.id + "/" + u.id]); }
    else if (r.v === "client") { html = vClient(r.a, r.b); var cb = findClient(r.a, r.b); c.push(["Business Operations", "#/k/biz"], ["Jama Star Rent a Car", "#/u/biz/jama"]); if (cb) c.push([cb.customer, location.hash]); }
    else if (r.v === "record") { var rec = findR(r.a); html = vRecord(rec); if (rec) { c.push([rec.k.name, "#/k/" + rec.k.id], [rec.u.name, "#/u/" + rec.k.id + "/" + rec.u.id], [rec.name, "#/record/" + rec.id]); } }
    else html = vHome();
    crumbs(c);
    $("main").innerHTML = html + '<p class="foot">ATLAS OS v' + VERSION + " · " + (S.mode === "live" ? "live from your OneDrive · read " + S.data.at.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }) : "developer preview · sample data") + "</p>";
  }

  // ---------------------------------------------------------------- actions
  var toastT;
  function toast(m) { var el = $("toast"); if (!el) { el = document.createElement("div"); el.id = "toast"; el.className = "toast"; el.setAttribute("role", "status"); document.body.appendChild(el); } el.textContent = m; el.hidden = false; clearTimeout(toastT); toastT = setTimeout(function () { el.hidden = true; }, 4200); }
  function openExcel(recId, note) {
    var r = findR(recId);
    var f = r && r.source ? S.files[r.source] : null;
    if (S.mode === "live" && f && f.webUrl) { window.open(f.webUrl, "_blank", "noopener"); }
    else toast((note ? note + " — " : "") + "Excel opens here once your Microsoft account is connected.");
  }
  function approve(id) {
    var x = inboxItems().filter(function (i) { return i.id === id; })[0]; if (!x) return;
    S.done[id] = new Date().toISOString(); store("done", S.done); logAct("Approved: " + x.title);
    openExcel(x.rec, "Approved · go to " + x.sheet); render();
  }
  function toggleTheme() { S.theme = S.theme === "dark" ? "light" : "dark"; store("theme", S.theme); render(); }

  document.addEventListener("click", function (e) {
    var el = e.target.closest("[data-go],[data-act],[data-approve],[data-later],[data-undo],[data-excel],[data-cmd]"); if (!el) return;
    var a = el.dataset;
    if (a.go) { if (location.hash === a.go) render(); else location.hash = a.go; closeCmd(); window.scrollTo(0, 0); }
    else if (a.act === "signin") signIn();
    else if (a.act === "reload") load();
    else if (a.approve) approve(a.approve);
    else if (a.later) toast("Kept in the Inbox for later.");
    else if (a.undo) { delete S.done[a.undo]; store("done", S.done); render(); }
    else if (a.excel) openExcel(a.excel);
    else if (a.cmd != null) runCmd(+a.cmd);
  });

  // ---------------------------------------------------------------- command bar
  var cmdSel = 0, cmdHits = [];
  function commands() {
    var c = [{ l: "CHIEF Home", g: "Go", go: "#/home" }, { l: "Morning Brief", g: "Go", go: "#/brief" }, { l: "Approval Inbox", g: "Go", go: "#/inbox" }, { l: "Record Registry", g: "Go", go: "#/registry" },
      { l: "Switch light / dark", g: "Action", fn: toggleTheme }, { l: "Connect Microsoft account", g: "Action", fn: signIn }, { l: "Refresh data", g: "Action", fn: load }];
    KINGDOMS.forEach(function (k) { c.push({ l: k.name, g: "Kingdom", go: "#/k/" + k.id }); k.units.forEach(function (u) { c.push({ l: u.name, g: k.name, go: "#/u/" + k.id + "/" + u.id }); }); });
    RECORDS.forEach(function (r) { c.push({ l: "Open " + r.name, s: r.id, g: "Record · " + r.id, go: "#/record/" + r.id }); });
    if (S.data) {
      allClients().forEach(function (x) { c.push({ l: "Account · " + x.customer, s: x.key + " " + x.name + " client account", g: "Jama · " + x.unit, go: "#/client/" + x.unit + "/" + encodeURIComponent(x.id) }); });
      S.data.kaec.forEach(function (v) { c.push({ l: "Issue invoice · " + v.name, s: v.key + " " + v.customer, g: v.customer, fn: function () { location.hash = "#/inbox"; openExcel("JAM-FIN-001", "Invoice " + v.name); } }); });
      S.data.wwc.forEach(function (a) { c.push({ l: "SOA reminder · " + a.customer, s: a.key + " " + a.plate + " soa statement", g: "Jama · WWC", fn: function () { openExcel("JAM-FIN-001", "SOA " + a.customer); } }); });
    }
    return c;
  }
  function filterCmd(q) {
    var words = q.toLowerCase().split(/\s+/).filter(Boolean), all = commands();
    cmdHits = !words.length ? all.slice(0, 8) : all.filter(function (c) { var s = (c.l + " " + (c.s || "") + " " + c.g).toLowerCase(); return words.every(function (w) { return s.indexOf(w) >= 0; }); }).slice(0, 12);
    cmdSel = 0; drawCmd();
  }
  function drawCmd() {
    $("cmdList").innerHTML = cmdHits.length ? cmdHits.map(function (c, i) { return '<li><button role="option" aria-selected="' + (i === cmdSel) + '" data-cmd="' + i + '">' + esc(c.l) + '<span class="g">' + esc(c.g) + "</span></button></li>"; }).join("") : '<li class="hint">No match. Try a kingdom, a record, a plate or a customer.</li>';
  }
  function runCmd(i) { var c = cmdHits[i]; if (!c) return; closeCmd(); if (c.go) location.hash = c.go; if (c.fn) c.fn(); }
  function openCmd() { $("scrim").classList.add("open"); $("cmdInput").value = ""; filterCmd(""); setTimeout(function () { $("cmdInput").focus(); }, 0); }
  function closeCmd() { $("scrim").classList.remove("open"); }
  $("cmdOpen").addEventListener("click", openCmd);
  $("scrim").addEventListener("click", function (e) { if (e.target === this) closeCmd(); });
  $("cmdInput").addEventListener("input", function () { filterCmd(this.value); });
  $("cmdInput").addEventListener("keydown", function (e) {
    if (e.key === "ArrowDown") { cmdSel = Math.min(cmdSel + 1, cmdHits.length - 1); drawCmd(); e.preventDefault(); }
    else if (e.key === "ArrowUp") { cmdSel = Math.max(cmdSel - 1, 0); drawCmd(); e.preventDefault(); }
    else if (e.key === "Enter") runCmd(cmdSel);
    else if (e.key === "Escape") closeCmd();
  });
  document.addEventListener("keydown", function (e) { if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "k") { e.preventDefault(); openCmd(); } });

  $("theme").addEventListener("click", toggleTheme);
  $("account").addEventListener("click", function () { if (S.account) toast("Signed in as " + (S.account.username || S.account.name)); else signIn(); });
  window.addEventListener("hashchange", render);
  render(); load();
})();
