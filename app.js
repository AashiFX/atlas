/* ATLAS OS v1.0.0 — by AashiFX
   CHIEF › Kingdom › Unit › Record. Morning Brief, Approval Inbox, Command Bar, Record Registry.
   Data: sample data until Microsoft sign-in is configured (config.js), then live from OneDrive via Microsoft Graph.
   This repository contains no business data. */
(function () {
  "use strict";
  var VERSION = "1.1.0";
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
        R("JAM-FIN-001", "XLS", "Jama Client Accounts", "Invoices, payments, SOA · KAEC + WWC", "LIVE", "jamaClientAccounts"),
        R("JAM-FIN-002", "XLS", "Jama VAT workbook", "Quarterly VAT IN / OUT", "TO CONNECT"),
        R("JAM-FIN-003", "XLS", "ADIF payment allocation", "Monthly allocation records", "TO CONNECT"),
        R("JAM-FLT-001", "XLS", "Vehicle Maintenance Record", "Fleet service history", "TO CONNECT"),
        R("JAM-FIN-004", "XLS", "Monthly Finance Report", "Month-end summary", "TO CONNECT")] },
      { id: "mas", name: "MAS Mobility", about: "Premium chauffeur & charter transport", records: [R("MAS-ADM-001", "XLS", "MAS records", "Point ATLAS to MAS workbooks", "TO CONNECT")] },
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
  var S = { theme: store("theme") || "dark", data: null, live: false, error: null, account: null, done: store("done") || {}, log: store("log") || [] };
  function logAct(text) { S.log.unshift({ t: new Date().toISOString(), text: text }); S.log = S.log.slice(0, 200); store("log", S.log); }

  // ---------------------------------------------------------------- data model (same rules as the Excel add-in)
  function isActive(v) {
    var s = String(v.status || "").toLowerCase();
    if (/closed|transfer|inactive/.test(s)) return false;
    if (v.customer && v.next && !/no current/i.test(v.customer)) return true;
    return /active/.test(s);
  }
  function model(raw) {
    function recs(rows, unit) {
      return (rows || []).filter(function (r) { return r[0] !== "" && r[0] != null; }).map(function (r) {
        return { unit: unit, key: String(r[0]).trim(), no: r[1], start: toDate(r[2]), end: toDate(r[3]), inv: r[4] || "", invStatus: r[5] || "", debit: num(r[6]), credit: num(r[7]), bal: num(r[8]), payStatus: r[11] || "" };
      });
    }
    var kRec = recs(raw.kRec, "KAEC"), wRec = recs(raw.wRec, "WWC");
    function rate(key) { var x = 0; kRec.forEach(function (c) { if (c.key === key && /issued/i.test(c.invStatus) && c.debit > 0) x = c.debit; }); return x; }
    var rows = [];
    for (var i = 0; i < raw.kSum.length; i++) { var a0 = String(raw.kSum[i][0] || "").trim(); if (/^Vehicle\b/.test(a0)) break; if (a0) rows.push(raw.kSum[i]); }
    var kaec = rows.map(function (r) {
      var key = String(r[0]).trim();
      return { unit: "KAEC", key: key, customer: r[1] || "", plate: r[2] || "", name: vName(r[2], key), out: num(r[8]), next: toDate(r[9]), status: String(r[10] || ""), rate: rate(key) };
    });
    var wwc = raw.att.filter(function (r) { return r[0] === "WWC"; }).map(function (r) {
      var key = String(r[1]).trim(), oldest = null;
      wRec.forEach(function (c) { if (c.key === key && c.bal > 0 && c.start && (!oldest || c.start < oldest)) oldest = c.start; });
      return { unit: "WWC", key: key, customer: r[2] || "", plate: r[3] || "", name: vName(r[3], key), rate: num(r[4]), next: toDate(r[6]), out: num(r[8]), oldest: oldest, daily: /^Daily /.test(key) };
    });
    return { kaec: kaec, active: kaec.filter(isActive), wwc: wwc, kRec: kRec, wRec: wRec, at: new Date(), webUrl: raw.webUrl || null };
  }
  function sample() {
    // Fictional sample data — real figures come only from your own workbook after sign-in.
    var t = todayUTC() / 864e5 + 25569;
    var v = [["Sample A 1001", "Sample Customer A", "Sample Sedan | 1001 AAA", 2400, 2], ["Sample B 1002", "Sample Customer B", "Sample SUV | 1002 BBB", 2400, 4],
      ["Sample C 1003", "Sample Customer C", "Sample Sedan | 1003 CCC", 2100, 4], ["Sample D 1004", "Sample Customer D", "Sample Crossover | 1004 DDD", 2500, 12], ["Sample E 1005", "Sample Customer E", "Sample Hatchback | 1005 EEE", 2200, 27]];
    var kSum = v.map(function (r) { return [r[0], r[1], r[2], "", "", "", r[3], r[3], 0, t + r[4], "Active", "", ""]; });
    var kRec = v.map(function (r) { return [r[0], 1, t + r[4] - 30, t + r[4], "SAMPLE", "Issued", r[3], r[3], 0, "", "", "PAID"]; });
    var att = [["WWC", "Sample W 2001", "Sample Account W", "Sample Sedan | 2001 WWW", 2645, "", t + 9, 9, 7935, ""], ["WWC", "Sample X 2002", "Sample Account X", "Sample SUV | 2002 XXX", 2645, "", t + 20, 20, 2645, ""]];
    var wRec = [["Sample W 2001", 1, t - 70, t - 40, "S1", "Issued", 2645, 0, 2645], ["Sample X 2002", 1, t - 10, t + 20, "S2", "Issued", 2645, 0, 2645]];
    return { kSum: kSum, kRec: kRec, att: att, wRec: wRec };
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
  function readJamaLive() {
    return token().then(function (tk) {
      var name = CFG.records.jamaClientAccounts;
      return graph("/me/drive/root/search(q='" + encodeURIComponent(name.replace(/'/g, "''")) + "')?$select=id,name,webUrl", tk).then(function (res) {
        var f = (res.value || []).filter(function (x) { return x.name === name; })[0];
        if (!f) throw new Error("“" + name + "” was not found in your OneDrive.");
        var base = "/me/drive/items/" + f.id + "/workbook/worksheets/";
        var rng = function (sheet, addr) { return graph(base + encodeURIComponent(sheet) + "/range(address='" + addr + "')?$select=values", tk).then(function (x) { return x.values; }); };
        var tbl = function (sheet, fallback) {
          return graph(base + encodeURIComponent(sheet) + "/tables?$select=name", tk).then(function (t) {
            var n = t.value && t.value[0] && t.value[0].name;
            return n ? graph("/me/drive/items/" + f.id + "/workbook/tables/" + encodeURIComponent(n) + "/dataBodyRange?$select=values", tk).then(function (x) { return x.values; }) : rng(sheet, fallback);
          });
        };
        return Promise.all([rng("KAEC Vehicles", "A7:M60"), rng("Unified Control Center", "A12:J40"), tbl("KAEC Vehicles", "A23:M217"), tbl("WWC Vehicles", "A31:M168")])
          .then(function (v) { return { kSum: v[0], att: v[1], kRec: v[2], wRec: v[3], webUrl: f.webUrl }; });
      });
    });
  }
  function signIn() {
    if (!CFG.msClientId) { toast("Microsoft sign-in is set up in the admin step (planned 11 AM). Until then ATLAS shows sample data."); return; }
    msReady().then(function (m) { m.loginRedirect({ scopes: CFG.scopes }); });
  }

  function load() {
    S.error = null;
    return msReady().then(function () {
      if (S.account) return readJamaLive().then(function (raw) { S.live = true; return raw; });
      S.live = false; return sample();
    }).then(function (raw) { S.data = model(raw); }).catch(function (e) { S.error = e.message || String(e); S.live = false; S.data = model(sample()); }).then(render);
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
    if (S.error) return '<div class="notice"><span><b class="bad">Couldn’t read your workbook:</b> ' + esc(S.error) + " Showing sample data.</span><button class=\"btn sm\" data-act=\"reload\">Try again</button></div>";
    if (S.live) return "";
    return '<div class="notice"><span>Sample data — figures become live after Microsoft sign-in.</span><button class="btn sm" data-act="signin">Connect Microsoft account</button></div>';
  }
  function statusTone(s) { return s === "LIVE" ? "good" : s === "TO CONNECT" ? "warn" : ""; }

  function vHome() {
    var d = S.data, hr = new Date().getHours(), act = d.active;
    var hi = hr < 12 ? "Good morning" : hr < 17 ? "Good afternoon" : "Good evening";
    var due = act.filter(function (v) { var n = daysTo(v.next); return n != null && n <= 7; }).length;
    var kOut = d.kaec.reduce(function (s, v) { return s + v.out; }, 0), wOut = d.wwc.reduce(function (s, v) { return s + v.out; }, 0);
    var h = head("CHIEF Home", hi + ", CHIEF", "Ten kingdoms under one command.", '<div class="row"><button class="btn" data-go="#/brief">Morning Brief</button><button class="btn primary" data-go="#/inbox">Approval Inbox · ' + openCount() + "</button></div>") + sourceNote();
    h += '<div class="kpis">' + kpi("Jama · bills due in 7 days", due, "KAEC contracts", due ? "warn" : "good") + kpi("Jama · KAEC outstanding", "SAR " + sar(kOut), act.length + " vehicles on contract", kOut > 0 ? "bad" : "good") +
      kpi("Jama · WWC outstanding", "SAR " + sar(wOut), d.wwc.length + " accounts", wOut > 0 ? "bad" : "good") + kpi("Records connected", RECORDS.filter(function (r) { return r.status === "LIVE"; }).length + " of " + RECORDS.length, "Record Registry") + "</div>";
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
      '<p class="foot">Approve opens the record in Excel at the right sheet and logs your decision. One-click automatic run (Office Script via Power Automate) arrives in phase 2.</p>';
  }
  function vRegistry() {
    var h = head("Record Registry", RECORDS.length + " records", "Every workbook and document, filed by kingdom with an Evidence ID. IDs are proposed — confirm or change.");
    h += '<section class="card full"><div class="scroll"><table><thead><tr><th>EVIDENCE ID</th><th>RECORD</th><th>KINGDOM › UNIT</th><th>TYPE</th><th>STATUS</th></tr></thead><tbody>';
    RECORDS.forEach(function (r) { h += '<tr><td class="strong mono">' + r.id + '</td><td><button class="btn sm" data-go="#/record/' + r.id + '">' + esc(r.name) + "</button></td><td>" + esc(r.k.name) + " › " + esc(r.u.name) + "</td><td>" + r.kind + '</td><td class="tag ' + statusTone(r.status) + '">' + r.status + "</td></tr>"; });
    return h + "</tbody></table></div></section>";
  }
  function vKingdom(k) {
    if (!k) return '<div class="empty">Kingdom not found.</div>';
    return head("Kingdom", k.name, esc(k.about)) + "<h2>" + esc(k.unitLabel || "Units") + "</h2>" + '<div class="grid">' + k.units.map(function (u) {
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
    var d = S.data, list = [];
    d.kaec.forEach(function (v) { if (v.customer && !/no current/i.test(v.customer)) list.push({ unit: "KAEC", key: v.key, customer: v.customer, name: v.name, out: v.out, next: v.next, status: isActive(v) ? "Active" : (v.status || "—"), active: isActive(v), rate: v.rate }); });
    d.wwc.forEach(function (a) { list.push({ unit: "WWC", key: a.key, customer: a.customer, name: a.name, out: a.out, next: a.next, status: a.out > 0 ? "Outstanding" : "Active", active: true, rate: a.rate, daily: a.daily, oldest: a.oldest }); });
    return list;
  }
  function clientAccounts(tab) {
    var list = allClients(), tabs = [["all", "All clients"], ["kaec", "KAEC"], ["wwc", "WWC"], ["due", "Outstanding"]];
    var f = list.filter(function (c) { return tab === "kaec" ? c.unit === "KAEC" : tab === "wwc" ? c.unit === "WWC" : tab === "due" ? c.out > 0 : true; })
      .sort(function (a, b) { return (b.active - a.active) || (b.out - a.out) || ((a.next || 9e15) - (b.next || 9e15)); });
    var h = '<section class="card full"><div class="card-h"><h2>Client accounts</h2><div class="row" role="tablist">' + tabs.map(function (t) {
      var n = list.filter(function (c) { return t[0] === "kaec" ? c.unit === "KAEC" : t[0] === "wwc" ? c.unit === "WWC" : t[0] === "due" ? c.out > 0 : true; }).length;
      return '<button role="tab" aria-selected="' + (t[0] === tab) + '" class="btn sm' + (t[0] === tab ? " primary" : "") + '" data-go="#/u/biz/jama/' + t[0] + '">' + t[1] + " · " + n + "</button>";
    }).join("") + '</div></div><div class="scroll"><table><thead><tr><th>CLIENT</th><th>VEHICLE</th><th>UNIT</th><th class="num">OUTSTANDING</th><th class="num">NEXT BILLING</th><th>STATUS</th></tr></thead><tbody>';
    f.forEach(function (c) {
      h += '<tr class="click" data-go="#/client/' + c.unit + "/" + encodeURIComponent(c.key) + '"><td class="strong">' + esc(c.customer) + "</td><td>" + esc(c.name) + "</td><td>" + c.unit + (c.daily ? " · Daily" : "") + '</td><td class="num mono ' + (c.out > 0 ? "bad" : "good") + '">' + sar(c.out) + '</td><td class="num">' + fmt(c.next) + '</td><td class="' + (c.out > 0 ? "bad" : c.active ? "good" : "") + '">' + esc(c.status) + "</td></tr>";
    });
    if (!f.length) h += '<tr><td colspan="6">No clients in this view.</td></tr>';
    return h + "</tbody></table></div></section>";
  }
  function vClient(unit, key) {
    var c = allClients().filter(function (x) { return x.unit === unit && x.key === key; })[0];
    if (!c) return '<div class="empty">Client account not found. <button class="btn sm" data-go="#/u/biz/jama">Back to Jama Star</button></div>';
    var cyc = (unit === "KAEC" ? S.data.kRec : S.data.wRec).filter(function (x) { return x.key === key; });
    var issued = cyc.filter(function (x) { return /issued/i.test(x.invStatus) && !/not/i.test(x.invStatus); });
    var billed = issued.reduce(function (s, x) { return s + x.debit; }, 0), recv = issued.reduce(function (s, x) { return s + x.credit; }, 0);
    var upcoming = cyc.filter(function (x) { return !/issued/i.test(x.invStatus) || /not/i.test(x.invStatus); }).slice(0, 2);
    var h = head("Jama Star › " + unit + " client", c.customer, esc(c.name) + " · " + unit + (c.daily ? " daily" : ""), '<div class="row"><button class="btn" data-excel="JAM-FIN-001">Statement (SOA)</button><button class="btn primary" data-excel="JAM-FIN-001">Open in Excel</button></div>') + sourceNote();
    h += '<div class="kpis">' + kpi("Monthly rate", c.rate ? "SAR " + sar(c.rate) : "—", unit) + kpi("Total billed", "SAR " + sar(billed), issued.length + " invoices") + kpi("Total received", "SAR " + sar(recv), "") +
      kpi("Outstanding", "SAR " + sar(c.out), c.out > 0 ? (c.oldest ? "Oldest unpaid " + fmt(c.oldest) : "Follow up") : "Fully collected", c.out > 0 ? "bad" : "good") + kpi("Next billing", fmt(c.next), c.next ? daysTo(c.next) + " days" : "", "warn") + "</div>";
    h += '<section class="card full"><div class="card-h"><h2>Account history</h2><small>' + issued.length + " issued · next " + upcoming.length + '</small></div><div class="scroll"><table><thead><tr><th>#</th><th>PERIOD</th><th>INVOICE</th><th class="num">DEBIT</th><th class="num">CREDIT</th><th class="num">BALANCE</th><th>STATUS</th></tr></thead><tbody>';
    issued.concat(upcoming).forEach(function (x) {
      var st = String(x.payStatus || x.invStatus), cls = /^paid/i.test(st) ? "good" : /unpaid|overdue/i.test(st) ? "bad" : /due/i.test(st) ? "warn" : "";
      h += "<tr><td>" + esc(x.no) + "</td><td>" + (x.start ? fmt(x.start) + " → " + fmt(x.end) : "—") + '</td><td class="strong">' + esc(x.inv || "—") + '</td><td class="num mono">' + (x.debit ? sar(x.debit) : "—") + '</td><td class="num mono">' + (x.credit ? sar(x.credit) : "—") + '</td><td class="num mono ' + (x.bal > 0 ? "bad" : "") + '">' + (x.debit ? sar(x.bal) : "—") + '</td><td class="strong ' + cls + '">' + esc(st) + "</td></tr>";
    });
    if (!cyc.length) h += '<tr><td colspan="7">History appears after Microsoft sign-in.</td></tr>';
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
      var d = S.data, act = d.active.slice().sort(function (a, b) { return a.next - b.next; });
      h += sourceNote() + jamaKpis() + '<section class="card full"><div class="card-h"><h2>KAEC · next billing</h2></div><div class="scroll"><table><thead><tr><th>VEHICLE</th><th>CUSTOMER</th><th class="num">SAR / MONTH</th><th class="num">NEXT BILLING</th></tr></thead><tbody>';
      act.forEach(function (v) { var n = daysTo(v.next); h += '<tr><td class="strong">' + esc(v.name) + "</td><td>" + esc(v.customer) + '</td><td class="num mono">' + sar(v.rate, 0) + '</td><td class="num strong ' + (n != null && n <= 7 ? "warn" : "") + '">' + fmtS(v.next) + "</td></tr>"; });
      h += '</tbody></table></div></section><section class="card full"><div class="card-h"><h2>WWC · outstanding</h2></div><div class="scroll"><table><thead><tr><th>CUSTOMER</th><th>OLDEST UNPAID</th><th class="num">SAR</th></tr></thead><tbody>';
      d.wwc.slice().sort(function (a, b) { return (a.oldest || 9e15) - (b.oldest || 9e15); }).forEach(function (a) { h += '<tr><td class="strong">' + esc(a.customer) + "</td><td>" + fmt(a.oldest) + '</td><td class="num bad mono">' + sar(a.out) + "</td></tr>"; });
      h += "</tbody></table></div></section>";
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
    if (!S.data) { $("main").innerHTML = '<div class="empty">Loading ATLAS…</div>'; return; }
    var k = r.a ? findK(r.a) : null, html, c = [["CHIEF", "#/home"]];
    if (r.v === "brief") { html = vBrief(); c.push(["Morning Brief", "#/brief"]); }
    else if (r.v === "inbox") { html = vInbox(); c.push(["Approval Inbox", "#/inbox"]); }
    else if (r.v === "registry") { html = vRegistry(); c.push(["Record Registry", "#/registry"]); }
    else if (r.v === "k") { html = vKingdom(k); if (k) c.push([k.name, "#/k/" + k.id]); }
    else if (r.v === "u") { var u = findU(k, r.b); html = vUnit(k, u); if (k) c.push([k.name, "#/k/" + k.id]); if (u) c.push([u.name, "#/u/" + k.id + "/" + u.id]); }
    else if (r.v === "client") { html = vClient(r.a, r.b); var cb = allClients().filter(function (x) { return x.unit === r.a && x.key === r.b; })[0]; c.push(["Business Operations", "#/k/biz"], ["Jama Star Rent a Car", "#/u/biz/jama"]); if (cb) c.push([cb.customer, location.hash]); }
    else if (r.v === "record") { var rec = findR(r.a); html = vRecord(rec); if (rec) { c.push([rec.k.name, "#/k/" + rec.k.id], [rec.u.name, "#/u/" + rec.k.id + "/" + rec.u.id], [rec.name, "#/record/" + rec.id]); } }
    else html = vHome();
    crumbs(c);
    $("main").innerHTML = html + '<p class="foot">ATLAS OS v' + VERSION + " · " + (S.live ? "live from your OneDrive · read " + S.data.at.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }) : "sample data") + "</p>";
  }

  // ---------------------------------------------------------------- actions
  var toastT;
  function toast(m) { var el = $("toast"); if (!el) { el = document.createElement("div"); el.id = "toast"; el.className = "toast"; el.setAttribute("role", "status"); document.body.appendChild(el); } el.textContent = m; el.hidden = false; clearTimeout(toastT); toastT = setTimeout(function () { el.hidden = true; }, 4200); }
  function openExcel(recId, note) {
    var r = findR(recId);
    if (S.live && S.data.webUrl && r && r.source === "jamaClientAccounts") { window.open(S.data.webUrl, "_blank", "noopener"); }
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
      allClients().forEach(function (x) { c.push({ l: "Account · " + x.customer, s: x.key + " " + x.name + " client account", g: "Jama · " + x.unit, go: "#/client/" + x.unit + "/" + encodeURIComponent(x.key) }); });
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
