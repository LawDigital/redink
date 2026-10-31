/* Dashboard data stays on this server. No third-party scripts or telemetry. */
"use strict";
(() => {
  const byId = (id) => document.getElementById(id);
  const numberFormat = new Intl.NumberFormat("en");
  const monthPattern = /^\d{4}-(0[1-9]|1[0-2])$/;
  const hostNames = { WD: "Word", EX: "Excel", OL: "Outlook" };
  const views = {
    overview: ["Overview", "License and add-in usage", "Daily add-in loads and durable monthly license evidence."],
    reports: ["License reports", "Monthly license reports", "Review distinct identities and export license reporting evidence."],
    logs: ["Function analytics", "Recorded function activity", "Recorded Red Ink function invocations by add-in, with run starts separated from tool operations."],
    configuration: ["Settings", "Server configuration", "Collection, publishing, security, and add-in configuration."],
    users: ["User statistics", "User activity statistics", "Observed features, frequency, and distribution by pseudonymous user."],
  };
  let requestController = null;
  let requestNumber = 0;
  let reportRangeInitialized = false;
  let followToday = true;
  let followCurrentMonth = true;
  let logRequest = 0;
  let logController = null;
  let showAllLogFunctions = false;
  let lastLogData = null;

  function element(tag, className, text) {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (text !== undefined) node.textContent = String(text);
    return node;
  }

  function formatNumber(value) {
    return typeof value === "number" && Number.isFinite(value) ? numberFormat.format(value) : "—";
  }

  function text(id, value) {
    byId(id).textContent = value === null || value === undefined || value === "" ? "—" : String(value);
  }

  function formatDate(value) {
    if (!value) return "No reports received";
    const date = new Date(value);
    return Number.isNaN(date.getTime()) ? "Unknown" : date.toLocaleString("en", { dateStyle: "medium", timeStyle: "short" });
  }

  function formatMonth(value, short = false) {
    if (!monthPattern.test(value || "")) return value || "—";
    const [year, month] = value.split("-").map(Number);
    const date = new Date(year, month - 1, 1);
    return date.toLocaleString("en", short ? { month: "short", year: "2-digit" } : { month: "long", year: "numeric" });
  }

  function localMonth() {
    const now = new Date();
    return now.getFullYear() + "-" + String(now.getMonth() + 1).padStart(2, "0");
  }

  function localDay() {
    return localMonth() + "-" + String(new Date().getDate()).padStart(2, "0");
  }

  function renderDaily(records, month) {
    const body = byId("daily-table");
    body.replaceChildren();
    for (const day of [...(records || [])].reverse()) {
      const row = element("tr");
      row.append(element("td", "", day.date), element("td", "numeric", formatNumber(day.users)));
      body.append(row);
    }
    if (!records || !records.length) emptyRow(body, 2, "No explicit daily load evidence has arrived for this month.");
    text("daily-coverage", formatNumber(month.users_with_daily_evidence) + " monthly identities have daily evidence; " + formatNumber(month.monthly_only_users) + " have monthly evidence only. Missing dates mean no evidence received, not confirmed zero use.");
  }

  function changeView() {
    const requested = location.hash.slice(1);
    const name = Object.prototype.hasOwnProperty.call(views, requested) ? requested : "overview";
    document.querySelectorAll("[data-panel]").forEach((node) => { node.hidden = node.dataset.panel !== name; });
    document.querySelectorAll("[data-view]").forEach((node) => {
      const active = node.dataset.view === name;
      node.classList.toggle("active", active);
      if (active) node.setAttribute("aria-current", "page");
      else node.removeAttribute("aria-current");
    });
    text("page-title", views[name][1]);
    text("page-description", views[name][2]);
    document.title = views[name][0] + " · Red Ink Usage Server";
    if (name === "logs") refreshLogs();
    if (name === "configuration") loadServerSettings();
    if (name === "users") refreshUserStats();
  }

  function setConnection(ok, message) {
    const node = byId("connection-status");
    node.classList.toggle("stale", !ok);
    node.replaceChildren(element("span", "status-dot"), document.createTextNode(message));
  }

  function emptyRow(body, columns, message) {
    const row = element("tr");
    const cell = element("td", "empty-row", message);
    cell.colSpan = columns;
    row.append(cell);
    body.append(row);
  }

  function renderHistory(history, month) {
    const records = Array.isArray(history) ? history.filter((row) => monthPattern.test(row.month || "")) : [];
    records.sort((a, b) => a.month.localeCompare(b.month));
    const body = byId("history-table");
    body.replaceChildren();
    for (const record of [...records].reverse()) {
      const row = element("tr");
      row.append(element("td", "", formatMonth(record.month)), element("td", "numeric", formatNumber(record.users)), element("td", "numeric", formatNumber(record.host_users)));
      body.append(row);
    }
    if (!records.length) emptyRow(body, 3, "No monthly usage evidence has arrived yet.");

    const chart = byId("history-chart");
    chart.replaceChildren();
    const visible = records.slice(-12);
    if (!visible.length) {
      chart.append(element("p", "empty-state", "Your monthly history will appear when reports arrive."));
      chart.setAttribute("aria-label", "No monthly usage evidence received.");
      return;
    }
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", "0 0 600 154");
    svg.setAttribute("aria-hidden", "true");
    const maximum = Math.max(1, ...visible.map(r => r.users || 0));
    const unit = 510 / visible.length;
    for (let i = 0; i <= 4; i++) {
      const y = 12 + i * 27;
      const line = document.createElementNS(svg.namespaceURI, "line");
      for (const [name, value] of Object.entries({x1: 45, x2: 592, y1: y, y2: y, stroke: "#353535", "stroke-width": .7})) line.setAttribute(name, value);
      svg.append(line);
    }
    for (const [i, record] of visible.entries()) {
      const count = Math.max(0, record.users || 0);
      const height = count / maximum * 102;
      const bar = document.createElementNS(svg.namespaceURI, "rect");
      for (const [name, value] of Object.entries({x: 52+i*unit, y: 120-height, width: Math.max(8, unit-9), height: Math.max(0,height), rx: 3, fill: record.month === month ? "#cf2e2e" : "#5b2f30"})) bar.setAttribute(name, value);
      svg.append(bar);
      const valueLabel = document.createElementNS(svg.namespaceURI, "text");
      valueLabel.textContent = String(count);
      for (const [name, value] of Object.entries({x: 52+i*unit+(unit-9)/2, y: Math.max(9, 114-height), "text-anchor": "middle", fill: "#f5f5f5", "font-size": 12, "font-weight": 700})) valueLabel.setAttribute(name, value);
      svg.append(valueLabel);
      const label = document.createElementNS(svg.namespaceURI, "text");
      label.textContent = record.month.slice(5);
      for (const [name,value] of Object.entries({x:52+i*unit+(unit-9)/2,y:139,"text-anchor":"middle",fill:"#a8a8a8","font-size":11})) label.setAttribute(name,value);
      svg.append(label);
    }
    chart.append(svg);
    chart.setAttribute("aria-label", visible.map(r => `${formatMonth(r.month)}: ${r.users} monthly license identities`).join("; "));
  }

  function renderHosts(hosts) {
    const box = byId("host-breakdown");
    box.replaceChildren();
    const counts = hosts && typeof hosts === "object" ? hosts : {};
    const maximum = Math.max(1, ...Object.values(counts).filter(Number.isFinite));
    for (const [code, name] of Object.entries(hostNames)) {
      const row = element("div", "host-row " + code.toLowerCase());
      const icon = element("span", "host-icon " + code.toLowerCase(), name[0]);
      icon.setAttribute("aria-hidden", "true");
      const label = element("div", "host-label", name);
      const track = element("div", "host-progress");
      const fill = element("span");
      fill.style.width = ((Number.isFinite(counts[code]) ? counts[code] : 0) / maximum * 100) + "%";
      track.append(fill);
      label.append(track);
      row.append(icon, label, element("span", "host-count", formatNumber(counts[code])));
      box.append(row);
    }
  }

  function renderRanking(id, records, key, transform = (value) => value) {
    const box = byId(id);
    box.replaceChildren();
    const rows = Array.isArray(records) ? records : [];
    if (!rows.length) {
      box.append(element("p", "empty-state", "No recognized entries yet."));
      return;
    }
    const maximum = Math.max(1, ...rows.map((row) => Number.isFinite(row.count) ? row.count : 0));
    for (const record of rows) {
      const row = element("div", "rank-row");
      const track = element("div", "rank-bar");
      const fill = element("span");
      fill.style.width = (Math.max(0, Number.isFinite(record.count) ? record.count : 0) / maximum * 100) + "%";
      track.append(fill);
      row.append(element("span", "rank-label", transform(record[key] || "Unknown")), element("span", "rank-count", formatNumber(record.count)), track);
      if (Array.isArray(record.commands)) {
        const detail=element("details", "group-commands"), summary=element("summary");
        const list=element("ul");
        for(const cmd of record.commands) list.append(element("li", "", `${cmd.feature} · ${formatNumber(cmd.count)}`));
        summary.append(row); detail.append(summary,list); box.append(detail);
        const tip=record.commands.map(c=>`${c.feature} · ${formatNumber(c.count)}`).join("\n");
        summary.title=tip; summary.setAttribute("aria-label", record.label+" — show commands");
      } else box.append(row);
    }
  }

  function renderLogs(logs) {
    text("log-files", formatNumber(logs.files));
    text("log-records", formatNumber(logs.records));
    text("log-errors", formatNumber(logs.errors));
    renderRanking("log-features", logs.by_feature, "feature");
    renderRanking("log-hosts", logs.by_host, "host", (host) => hostNames[host] || host);
    renderRanking("log-versions", logs.by_version, "version");
    const body = byId("log-days");
    body.replaceChildren();
    const days = Array.isArray(logs.by_day) ? logs.by_day : [];
    for (const day of days) {
      const row = element("tr");
      row.append(element("td", "", day.date || "Unknown"), element("td", "numeric", formatNumber(day.count)));
      body.append(row);
    }
    if (!days.length) emptyRow(body, 2, "No recognized entries yet.");
  }

  function renderLogTrend(days) {
    const target = byId("log-activity-chart");
    target.replaceChildren();
    if (!Array.isArray(days) || !days.length) {
      target.append(element("p", "empty-state", "No recorded invocations within the selected period. This is not proof of inactivity when log sources are missing."));
      target.setAttribute("aria-label", "No log evidence available for the selected range.");
      return;
    }
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", "0 0 600 154");
    svg.setAttribute("aria-hidden", "true");
    const maximum = Math.max(1, ...days.map(d => d.count || 0));
    const first = Date.parse(days[0].date + "T00:00:00Z");
    const last = Date.parse(days[days.length-1].date + "T00:00:00Z");
    const span = Math.max(86400000, last - first);
    const point = (d,value) => [48 + (Date.parse(d.date + "T00:00:00Z") - first) / span * 535, 122 - Math.max(0,value) / maximum * 99];
    const ns = "http://www.w3.org/2000/svg";
    const mk = (tag, attrs, txt) => { const node = document.createElementNS(ns,tag); for(const [k,v] of Object.entries(attrs)) node.setAttribute(k,v); if(txt !== undefined) node.textContent=txt; svg.append(node); return node; };
    for (let i = 0; i <= 4; i++) {
      const y = 22 + i*25;
      mk("line",{x1:48,y1:y,x2:585,y2:y,stroke:"#353535","stroke-width":.8});
      mk("text",{x:41,y:y+4,fill:"#a8a8a8","font-size":10,"text-anchor":"end"},String(Math.round(maximum*(4-i)/4)));
    }
    const series = [["standard","#5fa8ff"]];
    if (days.some(d=>d.agentic>0)) series.push(["agentic","#ff5a36"]);
    for(const [key,color] of series) {
      let segment=[];
      const flush=() => { if(segment.length===1 && days.length>=45) mk("circle",{cx:segment[0][0],cy:segment[0][1],r:2.6,fill:color}); if(segment.length>1) mk("polyline",{points:segment.map(p=>p.join(",")).join(" "),fill:"none",stroke:color,"stroke-width":2.5,"stroke-linecap":"round","stroke-linejoin":"round"}); segment=[]; };
      for(let i=0;i<days.length;i++) {
        // Do not connect points across dates without observed evidence.
        if(i>0 && Date.parse(days[i].date+"T00:00:00Z")-Date.parse(days[i-1].date+"T00:00:00Z")>86400000) flush();
        if(key === "agentic" && !days[i].agentic) { flush(); continue; }
        const p=point(days[i],days[i][key]||0);
        segment.push(p);
        if(days.length<45) mk("circle",{cx:p[0],cy:p[1],r:2.6,fill:color});
      }
      flush();
    }
    mk("text",{x:48,y:146,fill:"#a8a8a8","font-size":11},days[0].date);
    mk("text",{x:582,y:146,fill:"#a8a8a8","text-anchor":"end","font-size":11},days[days.length-1].date);
    target.append(svg);
    target.setAttribute("aria-label", days.map(d => `${d.date}: ${d.standard} other, ${d.agentic || "no evidenced"} top-level run starts`).join("; "));
  }

  function renderLogDetails(data) {
    const sourceRows = Array.isArray(data.source_status) ? data.source_status : [];
    text("log-invocations",formatNumber(data.invocations));
    text("log-agentic",data.agentic_invocations ? formatNumber(data.agentic_invocations) : "—");
    text("log-standard",formatNumber(data.standard_invocations));
    text("log-users",formatNumber(data.users));
    text("log-coverage",formatNumber(data.event_kinds?.internal || 0)+" internal calls excluded · "+formatNumber(data.files)+" imported files · "+formatNumber(data.scanned_sources)+" configured sources · "+formatNumber(data.other_entries)+" non-invocation log entries");
    text("log-last-event",data.last_entry_local ? data.last_entry_local.replace("T"," ")+" (client-local; no timezone)" : "No events observed in range");
    text("log-import-quality",sourceRows.length ? (formatNumber(data.import_errors)+" reported import issues") : "No log paths have been scanned. Configure [imports].log_paths.");
    const warning=byId("log-source-warning");
    if (!sourceRows.length) {
      warning.textContent="No log sources have been scanned. Set a readable LogPath source in Settings and save. Zero observed events does not imply zero usage.";
      warning.hidden=false;
    } else if(data.import_errors) {
      warning.textContent=`${data.import_errors} log import issues were reported. Previously saved records remain available; inspect source status below.`;
      warning.hidden=false;
    } else warning.hidden=true;
    text("log-activity-period",byId("log-range").selectedOptions[0].textContent);
    renderLogTrend(data.days);
    text("run-coverage",data.agentic_invocations ? "Only explicit run starts are shown; gaps are not evidence of inactivity." : "No explicit run starts were recorded in this selection. Older clients cannot establish the run count.");
    renderRanking("log-families",data.families || [],"label");
    renderRanking("log-detail-tools", (data.tool_details||[]).slice(0,40).map(r=>({...r,feature:r.kind+" · "+r.feature})), "feature");
    renderRanking("log-legacy",data.legacy_markers || [],"feature");
    renderRanking("log-agent-functions",data.agents || [],"feature");
    renderRanking("log-versions-detailed",(data.versions||[]).filter(r=>r.version).slice(0,15),"version");
    const hosts=byId("log-hosts-detailed");
    hosts.replaceChildren();
    for(const h of data.hosts || []) {
      const item=element("div","host-row");
      item.append(element("span","host-icon",(hostNames[h.host]||h.host)[0]));
      const details=element("div","host-label",hostNames[h.host]||h.host);
      details.append(element("p","metric-small",formatNumber(h.agentic)+" run starts / "+formatNumber(h.standard)+" other"));
      item.append(details,element("span","host-count",formatNumber(h.count)));
      hosts.append(item);
    }
    const body=byId("log-feature-table");
    body.replaceChildren();
    lastLogData=data;
    const listed=Array.isArray(data.features)?data.features:[];
    const visible=showAllLogFunctions?listed:listed.slice(0,25);
    text("log-feature-count",listed.length ? (`Showing ${visible.length} of ${listed.length} distinct function labels (up to 200)`):"No feature labels observed");
    const toggle=byId("log-show-more");
    toggle.hidden=listed.length<=25;
    toggle.textContent=showAllLogFunctions?"Show fewer":"Show all displayed functions";
    for(const item of visible) {
      const row=element("tr");
      const tag=element("span","category-tag "+(item.agentic?"agentic":""));
      tag.append(element("span","dot"),document.createTextNode(item.agentic?"Run start":"Command"));
      row.append(element("td","mono",item.feature),element("td","",item.group_label),element("td","",undefined),element("td","numeric",formatNumber(item.count)));
      row.children[2].append(tag);
      body.append(row);
    }
    if(!body.children.length) emptyRow(body,4,"No recognized feature invocations in this period.");
    const sources=byId("log-source-table");
    sources.replaceChildren();
    for(const source of sourceRows) {
      const row=element("tr");
      row.append(element("td","source-path",source.path),element("td","",formatDate(source.last_scan_utc)),element("td","numeric",formatNumber(source.records)),element("td","numeric",formatNumber(source.errors)),element("td","",source.detail||"No issues reported"));
      sources.append(row);
    }
    if(!sources.children.length) emptyRow(sources,5,"No log sources configured or scanned.");
  }

  async function refreshLogs() {
    if(location.hash !== "#logs") return;
    const id=++logRequest;
    if(logController) logController.abort();
    const controller=new AbortController();
    logController=controller;
    const timeout=setTimeout(()=>controller.abort(),20000);
    const query=new URLSearchParams({period:byId("log-range").value,host:byId("log-host").value});
    byId("download-log-analysis").href="/api/log-analysis.csv?"+query;
    try {
      const response=await fetch("/api/log-analysis?"+query,{credentials:"same-origin",cache:"no-store",signal:controller.signal,headers:{Accept:"application/json"}});
      if(response.status===401 || response.redirected){location.assign("/login");return;}
      if(!response.ok) throw new Error("HTTP "+response.status);
      const data=await response.json();
      if(id!==logRequest) return;
      renderLogDetails(data);
      byId("log-fetch-error").hidden=true;
    } catch(error) {
      if(id!==logRequest) return;
      const notice=byId("log-fetch-error");
      notice.textContent="Function analytics could not be refreshed. Previously displayed results may be stale. "+(error.name==="AbortError"?"Request timed out.":error.message);
      notice.hidden=false;
    } finally {clearTimeout(timeout);if(id===logRequest)logController=null;}
  }

  function renderConfiguration(files) {
    const box = byId("config-files");
    box.replaceChildren();
    const records = Array.isArray(files) ? files : [];
    for (const file of records) {
      const row = element("div", "configuration-file");
      const info = element("div", "file-info");
      info.append(element("strong", "", file.name || "Unknown file"));
      if (file.available) {
        const size = typeof file.bytes === "number" ? formatNumber(file.bytes) + " bytes" : "Size unavailable";
        info.append(element("p", "", size + (file.modified_utc ? " · Updated " + formatDate(file.modified_utc) : "")));
      } else info.append(element("p", "", "Place this file in the configured server directory to publish it."));
      if (file.url) {
        // Display only a URL supplied by this authenticated server, as plain text.
        // Do not turn arbitrary input into navigation or expose a secret token.
        try {
          const parsed = new URL(String(file.url), location.origin);
          if (parsed.origin === location.origin) info.append(element("p", "source-path", parsed.pathname));
        } catch (_) { /* An invalid optional URL is not displayed. */ }
      }
      row.append(element("span", "file-icon", "INI"), info, element("span", "pill " + (file.available ? "positive" : "neutral"), file.available ? "Available" : "Not published"));
      box.append(row);
    }
    if (!records.length) box.append(element("p", "empty-state", "Configuration publishing status is unavailable."));
  }

  function renderSources(sources) {
    const body = byId("sources-table");
    body.replaceChildren();
    const records = Array.isArray(sources) ? sources : [];
    for (const source of records) {
      const row = element("tr");
      const identity = element("td");
      identity.append(element("span", "source-kind", source.kind || "Import"), element("span", "source-path", source.path || "—"));
      row.append(identity, element("td", "", source.last_scan_utc ? formatDate(source.last_scan_utc) : "Not scanned"), element("td", "numeric", formatNumber(source.records)), element("td", "numeric", formatNumber(source.errors)));
      body.append(row);
    }
    if (!records.length) emptyRow(body, 4, "No file import sources configured. Web reports are collected independently.");
  }

  function updateExportLinks() {
    const from = byId("report-from").value;
    const to = byId("report-to").value;
    const valid = monthPattern.test(from) && monthPattern.test(to) && from <= to;
    const query = new URLSearchParams({ from, to }).toString();
    for (const [id, path] of [["download-monthly", "/api/reports/monthly.csv"], ["print-report", "/reports/print"]]) {
      const link = byId(id);
      if (valid) {
        link.href = path + "?" + query;
        link.removeAttribute("aria-disabled");
        link.removeAttribute("tabindex");
      } else {
        link.removeAttribute("href");
        link.setAttribute("aria-disabled", "true");
        link.setAttribute("tabindex", "-1");
      }
    }
    byId("report-range-error").hidden = valid;
    byId("download-daily").href = "/api/reports/daily.csv?" + new URLSearchParams({ month: byId("selected-month").value }).toString();
    byId("download-users").href = "/api/reports/users.csv?" + new URLSearchParams({ month: byId("selected-month").value }).toString();
  }

  function render(data) {
    const day = data.day || {};
    const month = data.month || {};
    const health = data.health || {};
    text("daily-users", formatNumber(day.users));
    text("monthly-users", formatNumber(month.users));
    text("monthly-user-days", formatNumber(month.observed_user_days));
    text("daily-date", day.date);
    text("monthly-period", formatMonth(month.month || byId("selected-month").value));
    text("last-ingest", formatDate(health.last_ingest_utc));
    text("database-status", health.database || "Status unavailable");
    text("pending-errors", formatNumber(health.pending_errors));
    text("last-refresh", formatDate(day.as_of || new Date().toISOString()));
    text("header-updated", formatDate(day.as_of || new Date().toISOString()));
    const healthLabel = byId("health-label");
    const healthy = ["ok", "healthy", "ready"].includes(String(health.status || "").toLowerCase());
    healthLabel.className = "pill " + (healthy ? "positive" : "warning");
    healthLabel.textContent = healthy ? "Operational" : (health.status || "Review status");
    const notes = byId("warnings");
    notes.replaceChildren();
    const warnings = Array.isArray(data.warnings) ? data.warnings : [];
    for (const warning of warnings) notes.append(element("li", "", warning));
    if (!warnings.length) {
      notes.append(element("li", "", "No collection warnings reported by the server."));
      notes.append(element("li", "", "Late monthly reports are assigned to their original usage month. A gap in received evidence is not proof of no usage."));
    }
    renderHistory(data.history, month.month || byId("selected-month").value);
    renderHosts(month.hosts);
    renderDaily(data.daily_history, month);
    // Log analytics has an independent date range and host filter.
    text("monthly-only-users", formatNumber(month.monthly_only_users));
    renderConfiguration(data.config_files);
    renderSources(data.sources);
    if (!reportRangeInitialized) {
      const history = Array.isArray(data.history) ? data.history : [];
      const months = history.map((row) => row.month).filter((value) => monthPattern.test(value || "")).sort();
      byId("report-from").value = months.length ? months[0] : byId("selected-month").value;
      byId("report-to").value = byId("selected-month").value;
      reportRangeInitialized = true;
    }
    updateExportLinks();
  }

  async function refresh() {
    // A dashboard left open overnight follows the new day unless the operator
    // deliberately selected a historical date/month.
    if (followToday) byId("selected-day").value = localDay();
    if (followCurrentMonth) byId("selected-month").value = localMonth();
    if (!monthPattern.test(byId("selected-month").value)) return;
    const current = ++requestNumber;
    if (requestController) requestController.abort();
    const controller = new AbortController();
    requestController = controller;
    const timeout = setTimeout(() => controller.abort(), 20000);
    byId("refresh").disabled = true;
    try {
      const response = await fetch("/api/dashboard?" + new URLSearchParams({ month: byId("selected-month").value, day: byId("selected-day").value }), { credentials: "same-origin", cache: "no-store", signal: controller.signal, headers: { Accept: "application/json" } });
      if (response.status === 401 || response.redirected) {
        location.assign("/login");
        return;
      }
      if (!response.ok) throw new Error("The server returned HTTP " + response.status + ".");
      const data = await response.json();
      if (current !== requestNumber) return;
      render(data);
      byId("fetch-error").hidden = true;
      setConnection(true, "Receiver connected");
    } catch (error) {
      if (current !== requestNumber) return;
      const notice = byId("fetch-error");
      notice.textContent = "Reporting data could not be refreshed. Previously displayed figures may be stale. " + (error.name === "AbortError" ? "The request timed out." : error.message);
      notice.hidden = false;
      setConnection(false, "Data may be stale");
    } finally {
      clearTimeout(timeout);
      if (current === requestNumber) { byId("refresh").disabled = false; requestController = null; }
    }
  }

  // Every action gives visible immediate press feedback, including navigation
  // and actions that later show their own success/error message.
  document.addEventListener("click",event=>{
    const control=event.target && event.target.closest ? event.target.closest("button, a.button, .nav-link") : null;
    if (!control) return;
    control.classList.remove("was-pressed");
    void control.offsetWidth;
    control.classList.add("was-pressed");
    setTimeout(()=>control.classList.remove("was-pressed"),1000);
  });

  /* Dedicated configuration and per-user views are not coupled to dashboard refreshes. */
  const csrf = document.querySelector('meta[name="csrf-token"]').content;
  const settingDescriptions = {
    server: {host:"Network bind address",port:"Listen port",threads:"HTTP worker threads",trusted_hosts:"Allowed browser hostnames (one per line)",secure_cookies:"Secure browser cookies (HTTPS)",allow_insecure_remote:"Allow unencrypted remote HTTP (not recommended)"},
    tls: {enabled:"Enable native HTTPS",certificate:"PEM certificate / chain path",private_key:"PEM private key path",key_password_env:"Environment variable for key password"},
    collection: {recent_report_window_seconds:"Legacy option (not used in current counting)",maximum_clock_skew_seconds:"Legacy option (not used in current counting)",maximum_request_bytes:"Maximum request bytes"},
    imports: {counter_paths:"LicenseCounter input folders or files (one path per line)",log_paths:"LogPath folders or files (one path per line)",recursive:"Scan subdirectories",scan_seconds:"Seconds between import scans",maximum_file_bytes:"Read budget per scan / file (bytes)",maximum_files_per_scan:"Files examined per scan"},
    publishing: {source_directory:"Alternate configuration publication folder (empty = default)"},
    billing: {unit_price:"Price per identity-month (empty = no price)",currency:"Billing currency (three letters)"},
    retention: {cutoff_date:"Earliest client-local date retained (YYYY-MM-DD; empty disables)",
      keep_months:"Rolling retention in calendar months (0 disables)",auto_cleanup:"Automatically delete expired evidence on scheduled scans"},
  };
  let settingsCache = null;
  let lastSettingsFetch = null;
  let selectedLogUser = null;
  function presentSettings(state) {
    settingsCache = state;
    const root = byId("settings-fields"); root.replaceChildren();
    const sections = state.sections || {};
    for (const [section, descriptions] of Object.entries(settingDescriptions)) {
      const panel = element("fieldset", "settings-section");
      panel.append(element("legend", "", section === "imports" ? "Import sources · LogPath" : section));
      for (const [key,label] of Object.entries(descriptions)) {
        const value = sections[section] ? sections[section][key] : undefined;
        if (value === undefined) continue;
        const row = element("div", "settings-field");
        const id = "setting-"+section+"-"+key;
        const caption = element("label", "", label); caption.htmlFor = id;
        const isArray = Array.isArray(value);
        const input = element(isArray ? "textarea" : "input");
        input.id = id;
        input.dataset.section = section;
        input.dataset.key = key;
        if (isArray) {input.rows = 3; input.value = value.join("\n");}
        else if (typeof value === "boolean") {input.type="checkbox"; input.checked=value;}
        else {input.type=typeof value === "number" ? "number" : "text";input.value=String(value);}
        row.append(caption, input);
        panel.append(row);
      }
      root.append(panel);
    }
    byId("passwordless-mode").checked = !!state.passwordless_local_only;
    byId("ingestion-token").value = state.tokens.ingestion || "";
    byId("config-token").value = state.tokens.config || "";
    byId("ingestion-token").placeholder = state.tokens.ingestion ? "" : "Previously generated token unavailable — rotate to display";
    byId("config-token").placeholder = state.tokens.config ? "" : "Previously generated token unavailable — rotate to display";
    text("default-publishing-dir", state.default_config_directory);
    text("active-publishing-dir", state.published_directory);
    updateAddinSnippet();
    updateClientUrls();
    refreshAutostart();
    text("settings-state", "Settings loaded · import and publishing paths apply on Save");
  }
  function updateAddinSnippet() {
    const token = byId("ingestion-token").value || "REPLACE_WITH_INGESTION_TOKEN";
    const configToken = byId("config-token").value || "REPLACE_WITH_CONFIG_TOKEN";
    // The browser's reachable origin is correct behind an HTTPS reverse
    // proxy; the internal listener may intentionally remain on 127.0.0.1.
    const base = location.origin;
    byId("addin-configuration").value =
      "LicenseCounterPath="+base+"/collect/"+token+"\nLicenseCounterMethod=WebExtended;RequireAck=true\nLicenseCounterIntervall=360\nLicenseCounterAnon=True\n"+
      "\n# Configuration download URLs (use in existing UpdateSource settings):\n"+
      "# " + base+"/config/"+configToken+"/redink.ini\n"+
      "# " + base+"/config/"+configToken+"/allmodels.ini\n"+
      "# " + base+"/config/"+configToken+"/specialservices.ini\n"+
      "# LogPath=<CLIENT_OR_NETWORK_LOG_FOLDER> (client parameter)";
  }
  async function copyClientValue(value, input, statusNode) {
    try {
      if (!value) throw new Error("Token is unavailable; rotate it in Settings first.");
      if (navigator.clipboard && window.isSecureContext) await navigator.clipboard.writeText(value);
      else { input.focus(); input.select(); if (!document.execCommand("copy")) throw new Error("Clipboard access denied; select the URL manually."); }
      statusNode.textContent="Copied to clipboard";
    } catch (error) {input.focus();input.select();statusNode.textContent=error.message+" Select and copy the highlighted URL.";}
  }
  function updateClientUrls() {
    const root=byId("client-url-list");
    root.replaceChildren();
    const configToken=byId("config-token").value;
    const ingestionToken=byId("ingestion-token").value;
    const names=["redink.ini","allmodels.ini","specialservices.ini","redink.ini.sig","allmodels.ini.sig","specialservices.ini.sig"];
    const links=[["LicenseCounterPath",ingestionToken,"/collect/"+encodeURIComponent(ingestionToken)],
      ...names.map(name=>[name,configToken,"/config/"+encodeURIComponent(configToken)+"/"+name])];
    for (const [label,token,path] of links) {
      const section=element("div","setting-row");
      const labelEl=element("label","",label);
      const input=element("input","mono");
      input.type="text";input.readOnly=true;input.autocomplete="off";
      input.value=token?new URL(path,location.origin).href:"";
      input.placeholder=token?"":"Token unavailable — rotate first";
      const copy=element("button","button secondary","Copy link");copy.type="button";copy.disabled=!token;
      const status=element("span","metric-small","");status.setAttribute("role","status");
      copy.addEventListener("click",()=>copyClientValue(input.value,input,status));
      section.append(labelEl,input,copy,status);root.append(section);
    }
  }
  async function refreshAutostart() {
    const status=byId("autostart-status"),enable=byId("autostart-enable"),disable=byId("autostart-disable");
    enable.disabled=true;disable.disabled=true;
    try {
      const result=await apiJson("/api/settings/autostart");
      status.textContent=result.enabled?"Enabled: starts headless at Windows logon":"Disabled: no current-user autostart task";
      enable.disabled=!!result.enabled;disable.disabled=!result.enabled;
    } catch(error){status.textContent="Autostart unavailable: "+error.message;}
  }
  for (const [buttonId,enabled] of [["autostart-enable",true],["autostart-disable",false]]) {
    byId(buttonId).addEventListener("click",async()=>{
      if (!confirm(enabled?"Enable automatic headless startup when you sign in to Windows?":"Disable automatic startup for this Windows user?"))return;
      const button=byId(buttonId);button.disabled=true;const original=button.textContent;button.textContent="Applying…";
      try {await apiJson("/api/settings/autostart",{method:"POST",body:JSON.stringify({enabled})});text("settings-state",enabled?"Windows autostart enabled":"Windows autostart disabled");await refreshAutostart();}
      catch(error){byId("autostart-status").textContent="Autostart change failed: "+error.message;}
      finally {button.textContent=original;}
    });
  }
  byId("shutdown-server").addEventListener("click",async()=>{
    if(!confirm("Shut down this Red Ink Usage Server now? This ends reporting and INI delivery until restarted."))return;
    const button=byId("shutdown-server");button.disabled=true;button.textContent="Shutting down…";
    try {await apiJson("/api/settings/shutdown",{method:"POST",body:"{}"});text("shutdown-state","Shutdown requested. This page will become unavailable; restart the server to use it again.");}
    catch(error){text("shutdown-state","Shutdown failed: "+error.message);button.disabled=false;button.textContent="Shut down server";}
  });
  async function apiJson(path, options={}) {
    const result = await fetch(path,{credentials:"same-origin",cache:"no-store",...options,
      headers:{Accept:"application/json",...(options.body?{"Content-Type":"application/json","X-CSRF-Token":csrf}:{}),...(options.headers||{})}});
    if(result.status===401 || result.redirected){location.assign("/login");throw new Error("Authentication required");}
    const value = await result.json();
    if(!result.ok) throw new Error(value.error || ("Request failed: "+result.status));
    return value;
  }
  async function loadServerSettings() {
    if (settingsCache || lastSettingsFetch) return;
    lastSettingsFetch = true;
    try {
      const value = await apiJson("/api/settings");
      presentSettings(value);
      byId("settings-error").hidden = true;
    } catch(e) {
      byId("settings-error").textContent = "Could not load settings: "+e.message;
      byId("settings-error").hidden = false;
    } finally {lastSettingsFetch = false;}
  }
  byId("save-settings").addEventListener("click", async () => {
    if(!settingsCache) return;
    const changes={};
    for(const input of byId("settings-fields").querySelectorAll("[data-section]")) {
      const section=input.dataset.section,key=input.dataset.key;
      const oldValue=settingsCache.sections[section][key];
      let next;
      if(Array.isArray(oldValue)) next=input.value.split(/\r?\n/).map(v=>v.trim()).filter(Boolean);
      else if(typeof oldValue==="boolean") next=input.checked;
      else if(typeof oldValue==="number") {next=Number(input.value);if(!Number.isSafeInteger(next)) {byId("settings-error").textContent="Invalid number: "+section+"."+key;byId("settings-error").hidden=false;return;}}
      else next=input.value;
      if(JSON.stringify(next)!==JSON.stringify(oldValue)) (changes[section] ||= {})[key] = next;
    }
    if(!Object.keys(changes).length){text("settings-state","No changes");return;}
    const saveButton=byId("save-settings"), oldLabel=saveButton.textContent;
    saveButton.disabled=true;
    saveButton.textContent="Saving…";
    text("settings-state","Validating and saving…");
    try {
      const outcome=await apiJson("/api/settings",{method:"POST",body:JSON.stringify(changes)});
      for(const [section,items] of Object.entries(changes)) Object.assign(settingsCache.sections[section],items);
      settingsCache.published_directory=outcome.published_directory || settingsCache.published_directory;
      text("active-publishing-dir",settingsCache.published_directory);
      byId("settings-error").hidden=!outcome.warnings?.length;
      if(outcome.warnings?.length) byId("settings-error").textContent=outcome.warnings.join(" · ");
      text("settings-state",outcome.restart_required ? "Saved · live fields applied · restart required for listener/TLS settings" :
           (outcome.scan_scheduled ? "Saved · sources updated · scan scheduled" : "Saved · changes active immediately"));
      updateAddinSnippet();
      if(Object.hasOwn(changes,"publishing")) refresh();
      if(Object.hasOwn(changes,"imports")) refreshLogs();
      if(Object.hasOwn(changes,"retention")) previewRetention();
      saveButton.textContent="✓ Saved";
    } catch(e) {byId("settings-error").textContent=e.message;byId("settings-error").hidden=false;saveButton.textContent="Save failed";text("settings-state","Save failed · no settings applied");}
    finally {saveButton.disabled=false;setTimeout(()=>{saveButton.textContent=oldLabel;},1800);}

  });
  async function previewRetention() {
    const area=byId("retention-preview"), button=byId("retention-clean");
    try {
      const data=await apiJson("/api/settings/retention/preview");
      area.replaceChildren();
      if (!data.enabled) {
        area.textContent="Retention is disabled. Configure a date or rolling period above, save, and restart.";
        button.disabled=true;return;
      }
      const items=Object.entries(data.counts||{});
      area.append(element("p","","Effective cutoff: "+data.cutoff_date+" · " + (data.total||0)+" stored records eligible"));
      const listing=element("ul","");
      for (const [key,value] of items) if(value) listing.append(element("li","",key.replace(/_/g," ")+": "+formatNumber(value)));
      if(!listing.children.length)listing.append(element("li","","No records currently require deletion."));
      area.append(listing);
      byId("retention-confirm").placeholder="DELETE BEFORE "+data.cutoff_date;
      button.dataset.cutoff=data.cutoff_date;
      button.disabled=!(data.total>0);
    } catch(error){area.textContent="Could not preview retention: "+error.message;button.disabled=true;}
  }
  byId("retention-preview-button").addEventListener("click",previewRetention);
  byId("retention-clean").addEventListener("click",async()=>{
    const button=byId("retention-clean"),cutoff=button.dataset.cutoff;
    const confirmation=byId("retention-confirm").value;
    if(confirmation!=="DELETE BEFORE "+cutoff) {byId("retention-preview").textContent="Enter DELETE BEFORE "+cutoff+" to confirm.";return;}
    if(!confirm("Permanently delete stored evidence before "+cutoff+"? Create an external database backup first."))return;
    try {
      const outcome=await apiJson("/api/settings/retention/clean",{method:"POST",body:JSON.stringify({confirmation})});
      text("settings-state","Retention cleaning completed · "+outcome.cutoff_date);
      byId("retention-confirm").value="";
      await previewRetention();
    }catch(error){byId("retention-preview").textContent="Cleanup failed: "+error.message;}
  });
  byId("save-passwordless").addEventListener("click",async () => {
    const enabled=byId("passwordless-mode").checked;
    if(enabled&&!confirm("Enable passwordless administration on this computer? Anyone with access to this localhost server session may administer the server."))return;
    try {const result=await apiJson("/api/settings/passwordless",{method:"POST",body:JSON.stringify({enabled})});text("settings-state",result.enabled?"Local passwordless sign-in enabled":"Password required");byId("settings-error").hidden=true;}
    catch(e){byId("settings-error").textContent=e.message;byId("settings-error").hidden=false;}
  });
  document.querySelectorAll("[data-reveal]").forEach(button=>button.addEventListener("click",()=>{const field=byId(button.dataset.reveal);field.type=field.type==="password"?"text":"password";}));
  document.querySelectorAll("[data-rotate]").forEach(button=>button.addEventListener("click",async () => {
    const kind=button.dataset.rotate;
    if(!confirm("Rotate the "+kind+" token now? The old token will immediately stop working. Update all add-in or deployment URLs that use it."))return;
    button.disabled=true;
    try {const data=await apiJson("/api/settings/token",{method:"POST",body:JSON.stringify({kind})});byId(kind+"-token").value=data.token;updateAddinSnippet();updateClientUrls();text("settings-state","Token rotated · update affected clients now");byId("settings-error").hidden=true;}
    catch(e){byId("settings-error").textContent=e.message;byId("settings-error").hidden=false;}
    finally{button.disabled=false;}
  }));
  byId("copy-addin-config").addEventListener("click",async()=>{try{await navigator.clipboard.writeText(byId("addin-configuration").value);text("settings-state","Add-in configuration copied");}catch(e){byId("addin-configuration").select();text("settings-state","Select and copy the configuration text");}});
  for(const [id,which] of [["open-publishing-default","default"],["open-publishing-active","published"]]) byId(id).addEventListener("click",async()=>{try{await apiJson("/api/settings/open-folder",{method:"POST",body:JSON.stringify({which})});text("settings-state","Folder opened on Windows server");}catch(e){byId("settings-error").textContent=e.message;byId("settings-error").hidden=false;}});
  async function refreshUserStats() {
    if(location.hash!=="#users")return;
    const query=new URLSearchParams({period:byId("user-period").value,host:byId("user-host").value});
    byId("user-export").href="/api/users/functions.csv?"+query;
    byId("license-users-export").href="/api/users/licenses.csv?"+new URLSearchParams({month:byId("selected-month").value});
    try {
      const [data,lic] = await Promise.all([apiJson("/api/users/functions?"+new URLSearchParams({...Object.fromEntries(query),...(selectedLogUser?{user:selectedLogUser}:{})})),apiJson("/api/users/licenses?"+new URLSearchParams({month:byId("selected-month").value}))]);
      byId("user-error").hidden=true;
      text("user-total",formatNumber(data.total_users));text("user-calls",formatNumber(data.total_invocations));
      const body=byId("user-body");body.replaceChildren();
      for(const user of data.users||[]){
        const row=element("tr");
        const button=element("button","user-select mono",user.user_hash);button.type="button";button.title="View this log user's activity";button.addEventListener("click",()=>{selectedLogUser=user.user_hash;refreshUserStats();});
        const cell=element("td");cell.append(button);
        row.append(cell,element("td","numeric",formatNumber(user.invocations)),element("td","numeric",formatNumber(user.active_days)),element("td","numeric",formatNumber(user.calls_per_active_day)),element("td","numeric",formatNumber(user.calls_per_calendar_day)),element("td","numeric",formatNumber(user.calls_per_7_days)),element("td","numeric",formatNumber(user.calls_per_active_week)),element("td","numeric",formatNumber(user.agentic)),element("td","",Object.keys(user.hosts).map(h=>hostNames[h]||h).join(", ")),element("td","numeric",formatNumber(user.distinct_features)));
        body.append(row);
      }
      if(!body.children.length)emptyRow(body,10,"No imported function events in this period.");
      renderRanking("user-distribution",(data.distribution||[]).map(r=>({name:r.range,count:r.users})),"name");
      const detail=byId("user-detail");detail.replaceChildren();
      if(data.selected){
        detail.append(element("p","mono",data.selected.user_hash),element("p","","Run starts: "+formatNumber(data.selected.agentic)+" · Other: "+formatNumber(data.selected.standard)));
        const u=data.selected;
        detail.append(element("p","",`Active weeks: ${u.active_weeks} · Days per active week: ${u.active_days_per_active_week} · Median per active day: ${u.median_per_active_day} · Peak daily activity: ${u.peak_day_count}`));
        detail.append(element("strong","","Weekly activity (week starting Monday)"));
        for(const week of (u.weeks||[]).slice(-12)) detail.append(element("p","user-feature",week.week_start+" · "+formatNumber(week.count)));
        if(u.tool_details?.length) detail.append(element("strong","","Most used tools inside runs (separate)"));
        for(const t of (u.tool_details||[]).slice(0,10)) detail.append(element("p","user-feature",t.feature+" · "+formatNumber(t.count)));
        const head=element("strong","","Most used functions");detail.append(head);
        for(const feature of (data.selected.features||[]).slice(0,12))detail.append(element("p","user-feature",feature.feature+" · "+formatNumber(feature.count)));
        detail.append(element("strong","","Observed days"));
        for(const day of (data.selected.days||[]).slice(-15))detail.append(element("p","user-feature",day.date+" · "+formatNumber(day.count)));
      } else detail.append(element("p","empty-state","Select a log user to view their observed feature mix and activity days."));
      const licenseBody=byId("license-user-body");licenseBody.replaceChildren();
      for(const user of lic.users||[]){const row=element("tr");row.append(element("td","mono",user.display_name||user.user_key),element("td","mono",user.network_scope||user.network_key),element("td","",user.product_id),element("td","",(user.hosts||"").split(",").map(h=>hostNames[h]||h).join(", ")),element("td","numeric",formatNumber(user.loaded_days)),element("td","",user.anonymous?"Pseudonymous":"Anon=False (name if available)"));licenseBody.append(row);}
      if(!licenseBody.children.length)emptyRow(licenseBody,6,"No monthly license identities received.");
    }catch(e){byId("user-error").textContent="Could not load user statistics. "+e.message;byId("user-error").hidden=false;}
  }
  byId("user-period").addEventListener("change",()=>{selectedLogUser=null;refreshUserStats();});
  byId("user-host").addEventListener("change",()=>{selectedLogUser=null;refreshUserStats();});
  byId("selected-day").value = localDay();
  byId("selected-month").value = localMonth();
  byId("report-from").value = localMonth();
  byId("report-to").value = localMonth();
  byId("selected-month").addEventListener("change", () => { followCurrentMonth = byId("selected-month").value === localMonth(); updateExportLinks(); refresh(); });
  byId("selected-day").addEventListener("change", () => {
    if (byId("selected-day").value) {
      followToday = byId("selected-day").value === localDay();
      byId("selected-month").value = byId("selected-day").value.slice(0, 7);
      followCurrentMonth = byId("selected-month").value === localMonth();
      updateExportLinks(); refresh();
    }
  });
  byId("report-from").addEventListener("change", updateExportLinks);
  byId("report-to").addEventListener("change", updateExportLinks);
  byId("refresh").addEventListener("click", () => { refresh(); refreshLogs(); });
  byId("log-range").addEventListener("change", refreshLogs);
  byId("log-host").addEventListener("change", refreshLogs);
  byId("log-show-more").addEventListener("click", () => {showAllLogFunctions = !showAllLogFunctions; if(lastLogData) renderLogDetails(lastLogData);});
  window.addEventListener("hashchange", changeView);
  document.addEventListener("visibilitychange", () => { if (!document.hidden) {refresh(); refreshLogs(); refreshUserStats();} });
  changeView();
  refresh();
  setInterval(() => { if (!document.hidden && !requestController) refresh(); if (!document.hidden && !logController) refreshLogs(); if (!document.hidden) refreshUserStats(); }, 30000);
})();
