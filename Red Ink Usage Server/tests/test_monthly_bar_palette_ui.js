/* Regression: chart data reds are legible, selected month distinctive, legend matches. */
"use strict";
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const assert = require("assert");
const root = path.join(__dirname, "../redink_server/static");
const source = fs.readFileSync(path.join(root, "app.js"), "utf8");
const css = fs.readFileSync(path.join(root, "app.css"), "utf8");
const start = source.indexOf("  function renderHistory(");
const end = source.indexOf("  function renderHosts(", start);
assert(start >= 0 && end > start);
class FakeNode {
  constructor(tag) {
    this.tag = tag;
    this.children = [];
    this.attrs = {};
    this.textContent = "";
    this.classes = new Set();
    this.classList = { add: (name) => this.classes.add(name) };
    this.namespaceURI = "http://www.w3.org/2000/svg";
  }
  append(...children) { this.children.push(...children); }
  replaceChildren(...children) { this.children = [...children]; }
  setAttribute(name, value) { this.attrs[name] = value; }
}
const chart = new FakeNode("div");
const table = new FakeNode("tbody");
const context = {
  byId: id => id === "history-table" ? table : id === "history-chart" ? chart : assert.fail(`unexpected ${id}`),
  monthPattern: /^\d{4}-(0[1-9]|1[0-2])$/,
  document: { createElementNS: (ns, name) => new FakeNode(name) },
  element: (tag, cls, content) => { const n = new FakeNode(tag); n.textContent = content || ""; return n; },
  formatMonth: s => s,
  formatNumber: n => n,
  emptyRow: () => assert.fail("unexpected empty row"),
};
vm.createContext(context);
vm.runInContext(source.slice(start, end), context);
context.renderHistory([{month:"2026-09",users:4,host_users:5},{month:"2026-10",users:6,host_users:8}],"2026-10");
const bars = chart.children[0].children.filter(n => n.tag === "rect");
assert.equal(bars.length, 2);
assert(bars[0].classes.has("monthly-bar"));
assert(!bars[0].classes.has("current"));
assert(bars[1].classes.has("monthly-bar"));
assert(bars[1].classes.has("current"));
assert(!bars.some(n => "fill" in n.attrs), "Chart fills must be owned by CSS, not hardcoded by JS");
assert(css.includes(".history-chart .monthly-bar{fill:#cf2e2e}"));
assert(css.includes(".history-chart .monthly-bar.current{fill:#ff5555}"));
assert(css.includes(".history-panel .legend i{background:#cf2e2e}"));
assert(css.includes("--brand-red:#b91312"));
console.log("Monthly-usage chart palette: PASS (past/current contrast, legend, semantic CSS, retained brand)");
