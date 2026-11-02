/* Optional Node-only regression test; browser UI has no npm dependency. */
"use strict";
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const assert = require("assert");
const source = fs.readFileSync(path.join(__dirname, "../redink_server/static/app.js"), "utf8");
function extract(start, end) {
  const a = source.indexOf(start);
  const b = source.indexOf(end, a + start.length);
  assert(a !== -1 && b !== -1, `Missing ${start}`);
  return source.slice(a, b);
}
class FakeElement {
  constructor(tag) { this.tag = tag; this.children = []; this.style = {}; this.attrs = {}; this.textContent = ""; }
  append(...children) { this.children.push(...children); }
  replaceChildren(...children) { this.children = [...children]; }
  setAttribute(name, value) { this.attrs[name] = value; }
}
const root = new FakeElement("div");
const context = {
  byId: (id) => { assert.equal(id, "log-versions-detailed"); return root; },
  hostNames: { WD: "Word", EX: "Excel", OL: "Outlook" },
  element: (tag, cls, value) => {
    const el = new FakeElement(tag);
    el.className = cls || "";
    if (value !== undefined) el.textContent = String(value);
    return el;
  },
  formatNumber: (value) => Number(value).toLocaleString("en-US"),
};
vm.createContext(context);
vm.runInContext(
  extract('  function renderRanking(', '  function renderLogs(') +
  extract('  function renderObservedAddinVersions(', '  function renderLogDetails('), context);
const show = (summary, rows) => context.renderObservedAddinVersions(summary, rows);
const sectionNames = () => root.children.map(s => s.children[0].children[0].textContent);
const summary = {users_by_addin:[{host:'WD',count:3},{host:'EX',count:1},{host:'OL',count:0}]};
show(summary, [{host:'EX',version:'V1',count:1},{host:'WD',version:'V3',count:2},
  {host:'WD',version:'V2',count:1},{host:'UNKNOWN',version:'bogus',count:5}]);
assert.equal(sectionNames().join(','), 'Word,Excel,Outlook');
assert.equal(root.children[0].children[0].children[1].textContent, '3 distinct users');
assert.equal(root.children[0].children[1].children.length, 2); // Two Word versions, no user double count.
assert.equal(root.children[0].children[1].children[0].children[0].textContent, 'V3');
assert.equal(root.children[0].children[1].children[0].children[1].textContent, '2');
assert.equal(root.children[1].children[1].children.length, 1);
assert.equal(root.children[2].children[1].children[0].textContent, 'No versions observed in this window.');
assert.equal(root.children[0].attrs['aria-label'], 'Word add-in versions');
show({users_by_addin:[{host:'OL',count:1}]},[{host:'OL',version:'Version unavailable',count:1}]);
assert.equal(sectionNames().join(','), 'Outlook');
assert.equal(root.children[0].children[1].children[0].children[0].textContent, 'Version unavailable');
show({users_by_addin:[]},[]);
assert.equal(root.children.length, 1);
assert.equal(root.children[0].textContent, 'No add-in versions observed in this window.');
console.log('Observed add-in versions grouping UI: PASS (all hosts, selected host, unknown version, empty window)');
