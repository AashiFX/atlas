/* ATLAS OS — Excel add-in shell v0.1.0 (Phase 1, read-only)
   AashiFX · Business Operations kingdom · Jama Star live from "Jama Client Accounts".
   This version only READS the workbook and opens sheets. It never writes cells. */
(function () {
  "use strict";

  var VERSION = "0.1.0";
  var S = {
    theme: "dark", unit: "jama", view: "dash", sel: null, pinned: "home",
    data: null, error: null, loading: false, live: false, query: ""
  };

  // ---------- Menu (kingdom: Business Operations › Jama Star) ----------
  var NAV = [
    { id: "home", label: "Dashboard", icon: "M3 11l9-7 9 7M5 10v10h14V10",
      subs: [{ label: "Overview", view: "dash" }, { label: "Control Center sheet", sheet: "Unified Control Center" }] },
    { id: "customers", label: "Customers", icon: "M16 20v-1a4 4 0 0 0-8 0v1M12 12a3.5 3.5 0 1 0 0-7 3.5 3.5 0 0 0 0 7",
      subs: [{ label: "WWC accounts", view: "wwc" }, { label: "KAEC customers", view: "kaec" }, { label: "SOA Center", sheet: "SOA Center" }] },
    { id: "vehicles", label: "Vehicles", icon: "M4 16h16M6 16l1.6-5h8.8L18 16M7 16v2.5M17 16v2.5",
      subs: [{ label: "KAEC fleet", view: "kaec" }, { label: "WWC fleet", view: "wwc" }, { label: "Vehicle master", sheet: "Master Data" }] },
    { id: "contracts", label: "Contracts", icon: "M7 3h7l4 4v14H7zM14 3v4h4M10 12h5M10 16h5",
      subs: [{ label: "Contract Actions", sheet: "Contract Actions" }, { label: "Contract Cycles", sheet: "Contract Cycles" }, { label: "Account Register", sheet: "Account Register" }] },
    { id: "billing", label: "Billing", icon: "M6 3h12v18l-3-2-3 2-3-2-3 2zM9 8h6M9 12h6",
      subs: [{ label: "Invoice Generator", sheet: "Invoice Generator" }, { label: "Daily Invoice", sheet: "Daily Invoice" }, { label: "Invoice Reprint", sheet: "Invoice Reprint" }, { label: "Invoice Record", sheet: "Invoice Record" }] },
    { id: "payments", label: "Payments", icon: "M3 7h18v10H3zM3 11h18M7 15h3",
      subs: [{ label: "Payment Entry", sheet: "Payment Entry" }, { label: "WWC Daily Payments", sheet: "WWC Daily Payments" }, { label: "Unified Ledger", sheet: "Unified Ledger" }] },
    { id: "actions", label: "Actions", icon: "M13 3L5 14h6l-1 7 8-11h-6z",
      subs: [{ label: "System Health", sheet: "System Health" }, { label: "Automation Log", sheet: "Automation Log" }, { label: "Refresh ATLAS data", act: "refresh" }] },
    { id: "targets", label: "Targets", icon: "M12 21a9 9 0 1 0 0-18 9 9 0 0 0 0 18M12 16a4 4 0 1 0 0-8 4 4 0 0 0 0 8",
      subs: [{ label: "Revenue targets", soon: true }, { label: "Utilisation", soon: true }, { label: "Collections", soon: true }] },
    { id: "reports", label: "Reports", icon: "M4 20V10M10 20V4M16 20v-7M22 20H2",
      subs: [{ label: "Accounts Report", sheet: "Accounts Report" }, { label: "KAEC Client SOA", sheet: "KAEC Client SOA" }, { label: "WWC Client SOA", sheet: "WWC Client SOA" }] },
    { id: "settings", label: "Settings", icon: "M12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6M12 2v3M12 19v3M2 12h3M19 12h3M4.9 4.9l2.1 2.1M17 17l2.1 2.1M4.9 19.1L7 17M17 7l2.1-2.1",
      subs: [{ label: "Light / Dark", act: "theme" }, { label: "Master Data", sheet: "Master Data" }, { label: "About ATLAS", view: "about" }] }
  ];
  var UNITS = [["jama", "Jama Star"], ["mas", "MAS"], ["prestige", "Prestige"]];

  // ---------- Helpers ----------
  function $(id) { return document.getElementById(id); }
  function esc(v) {
    return String(v == null ? "" : v).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  function num(v) { var n = typeof v === "number" ? v : parseFloat(String(v || "").replace(/,/g, "")); return isFinite(n) ? n : 0; }
  function sar(v, dp) { return num(v).toLocaleString("en-US", { minimumFractionDigits: dp == null ? 2 : dp, maximumFractionDigits: dp == null ? 2 : dp }); }
  var MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  function toDate(v) {
    if (v == null || v === "") return null;
    if (typeof v === "number" && v > 20000 && v < 80000) return new Date(Math.round((v - 25569) * 864e5));
    var s = String(v).trim(), m;
    if ((m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})/))) {
      var mi = MON.indexOf(m[2].charAt(0).toUpperCase() + m[2].slice(1, 3).toLowerCase());
      if (mi >= 0) return new Date(Date.UTC(+m[3], mi, +m[1]));
    }
    if ((m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/))) return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
    return null;
  }
  function fmt(d) { return d ? ("0" + d.getUTCDate()).slice(-2) + "-" + MON[d.getUTCMonth()] + "-" + d.getUTCFullYear() : "—"; }
  function fmtShort(d) { return d ? ("0" + d.getUTCDate()).slice(-2) + "-" + MON[d.getUTCMonth()] : "—"; }
  function todayUTC() { var n = new Date(); return Date.UTC(n.getFullYear(), n.getMonth(), n.getDate()); }
  function daysTo(d) { return d ? Math.round((d.getTime() - todayUTC()) / 864e5) : null; }
  function vehicleName(plate, key) {
    var p = String(plate || "").split("|").map(function (x) { return x.trim(); }).filter(Boolean);
    return p.length >= 2 ? p.slice(0, -1).join(" ") + " · " + p[p.length - 1] : (plate || key || "");
  }
  function store(k, v) { try { if (v === undefined) return localStorage.getItem(k); localStorage.setItem(k, v); } catch (e) { return null; } }
  function inExcel() { return typeof Office !== "undefined" && Office.context && Office.context.host === Office.HostType.Excel; }

  // ---------- Data ----------
  function readWorkbook() {
    return Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets;
      var kv = ws.getItemOrNullObject("KAEC Vehicles");
      var wv = ws.getItemOrNullObject("WWC Vehicles");
      var ucc = ws.getItemOrNullObject("Unified Control Center");
      return ctx.sync().then(function () {
        if (kv.isNullObject || ucc.isNullObject) throw new Error("Open the Jama Client Accounts workbook, then press Refresh.");
        kv.tables.load("items/name");
        if (!wv.isNullObject) wv.tables.load("items/name");
        return ctx.sync();
      }).then(function () {
        var out = {};
        out.kSum = kv.getRange("A7:M21"); out.kSum.load("values");
        out.att = ucc.getRange("A12:J40"); out.att.load("values");
        out.kRec = kv.tables.items.length ? kv.tables.items[0].getDataBodyRange() : kv.getRange("A23:M217");
        out.kRec.load("values");
        if (!wv.isNullObject) {
          out.wRec = wv.tables.items.length ? wv.tables.items[0].getDataBodyRange() : wv.getRange("A31:M168");
          out.wRec.load("values");
        }
        return ctx.sync().then(function () {
          return {
            kSum: out.kSum.values, att: out.att.values, kRec: out.kRec.values,
            wRec: out.wRec ? out.wRec.values : []
          };
        });
      });
    });
  }

  function model(raw) {
    var rec = function (rows, unit) {
      return rows.filter(function (r) { return r[0] !== "" && r[0] != null; }).map(function (r) {
        return { unit: unit, key: String(r[0]).trim(), no: r[1], start: toDate(r[2]), end: toDate(r[3]), inv: r[4] || "", invStatus: r[5] || "",
          debit: num(r[6]), credit: num(r[7]), bal: num(r[8]), paidOn: toDate(r[9]), method: r[10] || "", payStatus: r[11] || "" };
      });
    };
    var kRec = rec(raw.kRec, "KAEC"), wRec = rec(raw.wRec || [], "WWC");
    function lastRate(key, list) {
      var rate = 0;
      list.forEach(function (c) { if (c.key === key && String(c.invStatus).toLowerCase() === "issued" && c.debit > 0) rate = c.debit; });
      return rate;
    }
    var kaec = raw.kSum.filter(function (r) { return r[0] && String(r[0]).trim() !== "Vehicle"; }).map(function (r) {
      var key = String(r[0]).trim();
      return { unit: "KAEC", key: key, customer: r[1] || "", plate: r[2] || "", name: vehicleName(r[2], key), period: r[3] || "",
        lastInv: r[4] || "", lastDate: toDate(r[5]), billed: num(r[6]), received: num(r[7]), out: num(r[8]), next: toDate(r[9]),
        status: String(r[10] || ""), action: String(r[11] || ""), rate: lastRate(key, kRec) };
    });
    var att = raw.att.filter(function (r) { return r[0] === "WWC" || r[0] === "KAEC"; }).map(function (r) {
      var key = String(r[1]).trim();
      return { unit: r[0], key: key, customer: r[2] || "", plate: r[3] || "", name: vehicleName(r[3], key), rate: num(r[4]),
        lastInv: r[5] || "", next: toDate(r[6]), days: r[7], out: num(r[8]), action: String(r[9] || ""), daily: /^Daily /.test(key) };
    });
    var wwc = att.filter(function (a) { return a.unit === "WWC"; });
    // oldest unpaid WWC invoice per account
    wwc.forEach(function (a) {
      var oldest = null;
      wRec.forEach(function (c) { if (c.key === a.key && c.bal > 0 && c.start && (!oldest || c.start < oldest)) oldest = c.start; });
      a.oldest = oldest;
    });
    return { kaec: kaec, active: kaec.filter(function (v) { return v.status === "Active"; }), wwc: wwc, kRec: kRec, wRec: wRec, at: new Date() };
  }

  function demo() {
    // Fictional sample data for preview outside Excel. Real data is read live from the workbook only.
    var rows = [
      ["Sample A 1001", "Sample Customer A", "Sample Sedan | 1001 AAA", 2400, 3],
      ["Sample B 1002", "Sample Customer B", "Sample SUV | 1002 BBB", 2400, 5],
      ["Sample C 1003", "Sample Customer C", "Sample Sedan | 1003 CCC", 2100, 5],
      ["Sample D 1004", "Sample Customer D", "Sample Crossover | 1004 DDD", 2500, 13],
      ["Sample E 1005", "Sample Customer E", "Sample Hatchback | 1005 EEE", 2200, 28]
    ];
    var t = todayUTC() / 864e5 + 25569;
    var kSum = rows.map(function (r) { return [r[0], r[1], r[2], "", "", "", r[3], r[3], 0, t + r[4], "Active", "UPCOMING", "CURRENT CONTRACT"]; });
    var kRec = [];
    rows.forEach(function (r) {
      kRec.push([r[0], 1, t + r[4] - 30, t + r[4], "SAMPLE-INV-" + r[0].slice(-4), "Issued", r[3], r[3], 0, t + r[4] - 30, "Cash", "PAID", ""]);
      kRec.push([r[0], 2, t + r[4], t + r[4] + 30, "", "Not Issued", 0, 0, 0, "", "", "NOT DUE", ""]);
    });
    var att = [
      ["WWC", "Sample W 2001", "Sample Account W", "Sample Sedan | 2001 WWW", 2645, "SAMPLE-INV-2001", t + 10, 10, 5290, "Payment Follow-up"],
      ["WWC", "Sample X 2002", "Sample Account X", "Sample SUV | 2002 XXX", 2645, "SAMPLE-INV-2002", t + 20, 20, 2645, "Payment Follow-up"]
    ];
    return { kSum: kSum, att: att, kRec: kRec, wRec: [] };
  }

  function load() {
    S.loading = true; S.error = null; render();
    var p = inExcel() ? readWorkbook() : Promise.resolve(demo());
    return p.then(function (raw) { S.data = model(raw); S.live = inExcel(); })
      .catch(function (e) { S.error = (e && e.message) || String(e); })
      .then(function () { S.loading = false; render(); });
  }

  function openSheet(name) {
    if (!inExcel()) { toast("Preview mode: in Excel this opens the “" + name + "” sheet."); return; }
    Excel.run(function (ctx) {
      var s = ctx.workbook.worksheets.getItemOrNullObject(name);
      return ctx.sync().then(function () {
        if (s.isNullObject) { toast("Sheet “" + name + "” was not found in this workbook."); return; }
        s.activate(); return ctx.sync();
      });
    }).catch(function (e) { toast(e.message || String(e)); });
  }

  var toastTimer = null;
  function toast(msg) { S.toast = msg; render(); clearTimeout(toastTimer); toastTimer = setTimeout(function () { S.toast = null; render(); }, 4000); }

  // ---------- Render: shell ----------
  function icon(d, size) { return '<svg class="ic" aria-hidden="true" width="' + (size || 18) + '" height="' + (size || 18) + '" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><path d="' + d + '"/></svg>'; }

  function activeMenu() {
    return ({ dash: "home", wwc: "customers", kaec: "vehicles", vehicle: "vehicles", about: "settings", search: "home" })[S.view] || null;
  }

  function renderNav() {
    var act = activeMenu();
    var h = '<span class="nav-label">JAMA STAR</span>';
    NAV.forEach(function (n) {
      var open = S.pinned === n.id;
      h += '<div class="nav-item' + (open ? " open" : "") + '" data-menu="' + n.id + '">' +
        '<button class="nav-head' + (act === n.id ? " active" : "") + '" data-toggle="' + n.id + '" aria-expanded="' + open + '" title="' + esc(n.label) + '">' +
        icon(n.icon) + '<span class="lbl">' + esc(n.label) + '</span>' +
        '<span class="chev">' + icon(open ? "M6 9l6 6 6-6" : "M9 6l6 6-6 6", 14) + '</span></button>' +
        '<div class="subs" data-title="' + esc(n.label) + '">';
      n.subs.forEach(function (s, i) {
        h += '<button class="sub' + (s.soon ? " soon" : "") + '" data-sub="' + n.id + ":" + i + '">' + esc(s.label) + (s.soon ? " · soon" : "") + '</button>';
      });
      h += '</div></div>';
    });
    h += '<div class="nav-foot">Hover to peek · click to pin<br>ATLAS v' + VERSION + '</div>';
    $("nav").innerHTML = h;
  }

  function renderUnits() {
    $("units").innerHTML = UNITS.map(function (u) {
      return '<button data-unit="' + u[0] + '" aria-pressed="' + (S.unit === u[0]) + '">' + u[1] + '</button>';
    }).join("");
  }

  function render() {
    document.documentElement.setAttribute("data-theme", S.theme);
    var t = new Date();
    $("today").textContent = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][t.getDay()] + ", " + t.getDate() + " " + MON[t.getMonth()] + " " + t.getFullYear();
    renderUnits(); renderNav();
    var m = $("main"), h = "";
    if (S.toast) h += '<div class="banner" role="status">' + esc(S.toast) + '</div>';
    if (S.unit !== "jama") h += viewOtherUnit();
    else if (S.loading && !S.data) h += '<div class="empty">Reading Jama Client Accounts…</div>';
    else if (S.error) h += '<div class="empty"><h2>ATLAS could not read the workbook</h2><p>' + esc(S.error) + '</p><button class="btn primary" data-act="refresh">Try again</button></div>';
    else if (S.data) h += ({ dash: viewDash, kaec: viewKaec, wwc: viewWwc, vehicle: viewVehicle, about: viewAbout, search: viewSearch, soon: viewSoon })[S.view]();
    m.innerHTML = h;
  }

  function head(crumb, title, sub, right) {
    return '<div class="head"><div><div class="crumb">' + esc(crumb) + '</div><h1>' + esc(title) + '</h1>' +
      (sub ? '<div class="sub-title">' + sub + '</div>' : "") + '</div>' + (right || "") + '</div>';
  }
  function status() {
    return '<div class="status">' + (S.live ? "Live from this workbook" : "Preview data — open in Excel for live figures") +
      " · read " + (S.data ? S.data.at.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }) : "") + (S.loading ? " · refreshing…" : "") + '</div>';
  }

  // ---------- Views ----------
  function viewDash() {
    var d = S.data, act = d.active.slice().sort(function (a, b) { return (a.next || 0) - (b.next || 0); });
    var runRate = act.reduce(function (s, v) { return s + v.rate; }, 0);
    var kOut = d.kaec.reduce(function (s, v) { return s + v.out; }, 0);
    var wOut = d.wwc.reduce(function (s, v) { return s + v.out; }, 0);
    var wMonthly = d.wwc.filter(function (a) { return !a.daily; }).reduce(function (s, v) { return s + v.out; }, 0);
    var dueWeek = act.filter(function (v) { var n = daysTo(v.next); return n != null && n <= 7; });
    var title = dueWeek.length ? dueWeek.length + (dueWeek.length === 1 ? " bill" : " bills") + " due within 7 days" : "Everything in order";
    var h = head("BUSINESS OPERATIONS › JAMA STAR", title, null,
      '<div class="btns"><button class="btn" data-sheet="Payment Entry">Record payment</button><button class="btn" data-sheet="SOA Center">Generate SOA</button><button class="btn primary" data-sheet="Invoice Generator">Issue invoice</button></div>');
    h += '<div class="kpis">' +
      kpi("KAEC active vehicles", act.length, "All on contract", "good") +
      kpi("KAEC monthly run-rate", "SAR " + sar(runRate, 0), act.length + " contracts") +
      kpi("KAEC outstanding", "SAR " + sar(kOut), kOut > 0 ? "Follow up" : "Fully collected", kOut > 0 ? "bad" : "good") +
      kpi("WWC outstanding", "SAR " + sar(wOut), d.wwc.length + " accounts · SAR " + sar(wMonthly) + " monthly", wOut > 0 ? "bad" : "good") +
      '</div>';
    h += '<div class="split"><section class="card"><div class="card-h"><h2>KAEC · next billing</h2><small>Click a vehicle to open it</small></div><div class="scroll"><table><thead><tr><th>VEHICLE</th><th class="hide-n">CUSTOMER</th><th class="num">SAR / MONTH</th><th class="num">NEXT BILLING</th></tr></thead><tbody>';
    act.forEach(function (v) {
      var n = daysTo(v.next), cls = n == null ? "" : n <= 7 ? "warn" : "";
      h += '<tr class="click" data-vehicle="KAEC|' + esc(v.key) + '"><td class="strong">' + esc(v.name) + '</td><td class="hide-n">' + esc(v.customer) + '</td><td class="num">' + sar(v.rate, 0) + '</td><td class="num strong ' + cls + '">' + fmtShort(v.next) + '</td></tr>';
    });
    h += '</tbody></table></div></section><aside class="side"><h2>ATLAS suggests</h2>' + suggestions(dueWeek, kOut, wOut) + '</aside></div>' + status();
    return h;
  }
  function kpi(label, value, note, cls) {
    return '<div class="kpi"><span>' + esc(label) + '</span><b class="' + (cls === "bad" ? "bad" : cls === "good" && /SAR/.test(String(value)) ? "good" : "") + (typeof value === "string" && !/^(SAR|\d)/.test(value) ? " txt" : "") + '">' + esc(value) + '</b><small class="' + (cls || "") + '">' + esc(note) + '</small></div>';
  }
  function suggestions(dueWeek, kOut, wOut) {
    var h = "", d = S.data;
    var byDate = {};
    dueWeek.forEach(function (v) { var k = fmt(v.next); (byDate[k] = byDate[k] || []).push(v); });
    Object.keys(byDate).slice(0, 3).forEach(function (k) {
      var list = byDate[k], n = daysTo(list[0].next);
      var tag = n < 0 ? "OVERDUE " + (-n) + "D" : n === 0 ? "DUE TODAY" : "DUE IN " + n + (n === 1 ? " DAY" : " DAYS");
      h += '<div class="tip"><span class="tag ' + (n < 0 ? "bad" : "warn") + '">' + tag + '</span><b>' + esc(list.map(function (v) { return v.name; }).join(" & ")) + '</b><span>Billing ' + fmtShort(list[0].next) + ' · SAR ' + sar(list.reduce(function (s, v) { return s + v.rate; }, 0), 0) + (list.length > 1 ? " together" : "") + '</span></div>';
    });
    if (wOut > 0) {
      var old = d.wwc.filter(function (a) { return a.oldest; }).sort(function (a, b) { return a.oldest - b.oldest; }).slice(0, 2)
        .map(function (a) { return a.customer.replace(/^Mr\.?\s*/, "").split(" ")[0]; });
      h += '<div class="tip alert"><span class="tag bad">COLLECTIONS</span><b>WWC · SAR ' + sar(wOut) + ' open</b><span>' + (old.length ? "Oldest: " + esc(old.join(", ")) + " — send SOA reminder" : "Send SOA reminders") + '</span></div>';
    }
    if (kOut > 0) h += '<div class="tip alert"><span class="tag bad">KAEC</span><b>SAR ' + sar(kOut) + ' outstanding</b><span>Check Payment Entry</span></div>';
    if (!h) h = '<div class="tip"><span class="tag good">ALL CLEAR</span><b>Nothing urgent</b><span>No bills in the next 7 days, nothing outstanding.</span></div>';
    return h;
  }

  function viewKaec() {
    var list = S.data.kaec.slice().sort(function (a, b) { return (a.status !== "Active") - (b.status !== "Active") || (a.next || 0) - (b.next || 0); });
    var h = head("VEHICLES › KAEC", "KAEC fleet", list.length + " vehicle records · " + S.data.active.length + " active");
    h += '<section class="card" style="flex:none"><div class="scroll"><table><thead><tr><th>VEHICLE</th><th class="hide-n">CUSTOMER</th><th>STATUS</th><th class="num">OUTSTANDING</th><th class="num">NEXT BILLING</th></tr></thead><tbody>';
    list.forEach(function (v) {
      h += '<tr class="click" data-vehicle="KAEC|' + esc(v.key) + '"><td class="strong">' + esc(v.name) + '</td><td class="hide-n">' + esc(v.customer || "—") + '</td><td class="' + (v.status === "Active" ? "good" : "") + '">' + esc(v.status) + '</td><td class="num ' + (v.out > 0 ? "bad" : "") + '">' + sar(v.out) + '</td><td class="num">' + fmt(v.next) + '</td></tr>';
    });
    return h + '</tbody></table></div></section>' + status();
  }

  function viewWwc() {
    var list = S.data.wwc.slice().sort(function (a, b) { return (a.oldest || a.next || 9e15) - (b.oldest || b.next || 9e15); });
    var total = list.reduce(function (s, v) { return s + v.out; }, 0);
    var h = head("CUSTOMERS › WWC", "WWC accounts", list.length + " accounts · SAR " + sar(total) + " outstanding", '<div class="btns"><button class="btn" data-sheet="WWC Client SOA">WWC Client SOA</button></div>');
    h += '<section class="card" style="flex:none"><div class="scroll"><table><thead><tr><th>CUSTOMER</th><th class="hide-n">VEHICLE</th><th class="hide-n">OLDEST UNPAID</th><th class="num">OUTSTANDING</th><th class="num">NEXT BILLING</th></tr></thead><tbody>';
    list.forEach(function (a) {
      h += '<tr class="click" data-vehicle="WWC|' + esc(a.key) + '"><td class="strong">' + esc(a.customer) + '</td><td class="hide-n">' + esc(a.name) + '</td><td class="hide-n">' + fmt(a.oldest) + '</td><td class="num bad">' + sar(a.out) + '</td><td class="num">' + fmt(a.next) + '</td></tr>';
    });
    return h + '</tbody></table></div></section>' + status();
  }

  function viewVehicle() {
    var p = (S.sel || "").split("|"), unit = p[0], key = p.slice(1).join("|");
    var d = S.data, v = unit === "KAEC" ? d.kaec.filter(function (x) { return x.key === key; })[0] : d.wwc.filter(function (x) { return x.key === key; })[0];
    if (!v) return '<div class="empty">Vehicle not found. <button class="link" data-view="dash">Back to dashboard</button></div>';
    var cycles = (unit === "KAEC" ? d.kRec : d.wRec).filter(function (c) { return c.key === key; });
    var issued = cycles.filter(function (c) { return String(c.invStatus).toLowerCase() === "issued"; });
    var upcoming = cycles.filter(function (c) { return String(c.invStatus).toLowerCase() !== "issued"; }).slice(0, 2);
    var show = issued.slice(-4).concat(upcoming);
    var out = unit === "KAEC" ? v.out : v.out, rate = v.rate || (issued.length ? issued[issued.length - 1].debit : 0);
    var paid = out <= 0;
    var h = '<button class="link" data-view="' + (unit === "KAEC" ? "dash" : "wwc") + '">‹ Back</button>';
    h += head("VEHICLES › " + unit, v.name, esc(v.customer) + (unit === "KAEC" && v.status ? " · " + esc(v.status) : ""),
      '<span class="pill ' + (paid ? "good" : "bad") + '">' + (paid ? "PAID · up to date" : "SAR " + sar(out) + " open") + '</span>');
    h += '<div class="kpis">' + kpi("Customer", v.customer || "—", unit) + kpi("Monthly rate", "SAR " + sar(rate), "Last issued invoice") +
      kpi("Outstanding", "SAR " + sar(out), paid ? "Fully collected" : "Follow up", paid ? "good" : "bad") + kpi("Next billing", fmt(v.next), v.next ? (daysTo(v.next) + " days") : "", "warn") + '</div>';
    h += '<div class="split"><section class="card"><div class="card-h"><h2>Billing cycles</h2><small>Last 4 issued + next 2</small></div><div class="scroll"><table><thead><tr><th class="hide-n">#</th><th>PERIOD</th><th class="hide-n">INVOICE</th><th class="num">SAR</th><th class="num">STATUS</th></tr></thead><tbody>';
    show.forEach(function (c) {
      var st = String(c.payStatus || c.invStatus), cls = /^paid/i.test(st) ? "good" : /unpaid|overdue/i.test(st) ? "bad" : /due/i.test(st) ? "warn" : "";
      h += '<tr><td class="hide-n">' + esc(c.no) + '</td><td>' + (c.start ? fmt(c.start) + " → " + fmt(c.end) : "—") + '</td><td class="strong hide-n">' + esc(c.inv || "—") + '</td><td class="num">' + (c.debit ? sar(c.debit) : "—") + '</td><td class="num strong ' + cls + '">' + esc(st) + '</td></tr>';
    });
    if (!show.length) h += '<tr><td colspan="5">No cycles found for this vehicle.</td></tr>';
    h += '</tbody></table></div></section><aside class="side"><h2>Actions</h2>' +
      '<button class="btn primary block" data-sheet="Invoice Generator">Issue next invoice</button>' +
      '<button class="btn block" data-sheet="Payment Entry">Record payment</button>' +
      '<button class="btn block" data-sheet="' + (unit === "KAEC" ? "KAEC Client SOA" : "WWC Client SOA") + '">Statement of account</button>' +
      '<button class="btn block" data-sheet="' + esc(key) + '">Open vehicle sheet</button></aside></div>' + status();
    return h;
  }

  function viewSearch() {
    var q = S.query.toLowerCase(), d = S.data;
    var hits = d.kaec.concat(d.wwc).filter(function (v) { return (v.key + " " + v.customer + " " + v.plate + " " + v.lastInv).toLowerCase().indexOf(q) >= 0; });
    var h = head("SEARCH", "“" + S.query + "”", hits.length + " match" + (hits.length === 1 ? "" : "es"));
    if (!hits.length) return h + '<div class="empty">Nothing found. Try a plate number, a name or an invoice number.</div>';
    h += '<section class="card" style="flex:none"><div class="scroll"><table><thead><tr><th>UNIT</th><th>VEHICLE</th><th>CUSTOMER</th><th class="num">OUTSTANDING</th></tr></thead><tbody>';
    hits.forEach(function (v) { h += '<tr class="click" data-vehicle="' + v.unit + "|" + esc(v.key) + '"><td>' + v.unit + '</td><td class="strong">' + esc(v.name) + '</td><td>' + esc(v.customer) + '</td><td class="num">' + sar(v.out) + '</td></tr>'; });
    return h + '</tbody></table></div></section>';
  }

  function viewOtherUnit() {
    var n = S.unit === "mas" ? "MAS" : "Prestige";
    return head("BUSINESS OPERATIONS › " + n.toUpperCase(), n + " workspace") + '<div class="empty">' + n + ' joins ATLAS in a later roadmap phase. Jama Star is live first.</div>';
  }
  function viewSoon() { return head("ATLAS", S.soonLabel || "Coming soon") + '<div class="empty">This screen arrives in a later ATLAS phase.</div>'; }
  function viewAbout() {
    return head("SETTINGS", "About ATLAS OS") + '<div class="tip"><b>ATLAS OS v' + VERSION + ' · Phase 1</b><span>AashiFX · Business Operations kingdom · Jama Star live.</span><span>This version reads the workbook and opens sheets. It never changes your data.</span><span>Look: Steel Command (locked 18-Oct-2026). Theme: ' + S.theme + '.</span></div>';
  }

  // ---------- Events ----------
  function go(view, sel) { S.view = view; if (sel !== undefined) S.sel = sel; S.unit = "jama"; render(); var m = $("main"); if (m.scrollIntoView) window.scrollTo(0, 0); }

  function runSub(s) {
    if (s.soon) { S.soonLabel = s.label; go("soon"); }
    else if (s.view) go(s.view);
    else if (s.sheet) openSheet(s.sheet);
    else if (s.act === "theme") toggleTheme();
    else if (s.act === "refresh") load();
  }
  function toggleTheme() { S.theme = S.theme === "dark" ? "light" : "dark"; store("atlas.theme", S.theme); render(); }

  document.addEventListener("click", function (e) {
    var el = e.target.closest("[data-toggle],[data-sub],[data-unit],[data-vehicle],[data-sheet],[data-view],[data-act]");
    if (!el) { if (window.innerWidth <= 700 && S.pinned && !e.target.closest(".nav")) { S.pinned = null; renderNav(); } return; }
    var a = el.dataset;
    if (a.toggle) { S.pinned = S.pinned === a.toggle ? null : a.toggle; renderNav(); }
    else if (a.sub) { var p = a.sub.split(":"), n = NAV.filter(function (x) { return x.id === p[0]; })[0]; if (window.innerWidth <= 700) S.pinned = null; runSub(n.subs[+p[1]]); }
    else if (a.unit) { S.unit = a.unit; render(); }
    else if (a.vehicle) go("vehicle", a.vehicle);
    else if (a.sheet) openSheet(a.sheet);
    else if (a.view) go(a.view);
    else if (a.act === "refresh") load();
  });
  // hover-to-peek (wide layout): opens a menu while the pointer is over it
  document.addEventListener("mouseover", function (e) {
    var item = e.target.closest && e.target.closest(".nav-item");
    document.querySelectorAll(".nav-item.peek").forEach(function (x) { if (x !== item) x.classList.remove("peek", "open"); });
    if (item && !item.classList.contains("open")) item.classList.add("open", "peek");
  });

  function init() {
    S.theme = store("atlas.theme") || "dark";
    if (window.innerWidth <= 700) S.pinned = null;
    $("theme").addEventListener("click", toggleTheme);
    $("refresh").addEventListener("click", load);
    $("cmd").addEventListener("keydown", function (e) {
      if (e.key === "Enter" && this.value.trim()) { S.query = this.value.trim(); go("search"); }
    });
    load();
  }

  if (typeof Office !== "undefined" && Office.onReady) Office.onReady(function () { init(); });
  else document.addEventListener("DOMContentLoaded", init);
})();
