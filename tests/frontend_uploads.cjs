// Run the real page handlers with native File/FormData and a minimal DOM harness.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const {File} = require('node:buffer');

const elements = new Map();
function element(selector) {
  if (!elements.has(selector)) {
    const el = {innerHTML: '点击选择文件', files: [], clicks: 0,
      classList: {toggle() {}}, click() { this.clicks++; }};
    Object.defineProperty(el, 'value', {
      get() { return this._value || ''; },
      set(value) { this._value = value; if (value === '') this.files = []; },
    });
    elements.set(selector, el);
  }
  return elements.get(selector);
}
global.document = {querySelector: element, querySelectorAll: () => []};
let submitted;
global.fetch = async (url, options) => ({ok: true, json: async () => {
  if (url === '/api/dashboard') return {shipments: []};
  assert.equal(url, '/api/jobs'); submitted = options.body; return {id: 'test-job'};
}});
const html = fs.readFileSync(process.argv[2], 'utf8');
vm.runInThisContext(html.slice(html.indexOf('<script>') + 8, html.lastIndexOf('</script>'))
  .replace("loadHealth().then(() => go('dash'));", ''));
vm.runInThisContext("go = () => {}; health = {sample_template: true};");

function choose(selector, files) {
  const input = element(selector);
  input.files = files;
  input._value = files.length ? 'C:\\fakepath\\' + files[0].name : '';
  input.onchange();
}
function drop(selector, files) {
  element(selector).ondrop({preventDefault() {}, dataTransfer: {files}});
}
function remove(selector, index) {
  const button = {dataset: {removeFile: String(index)}, closest: () => button};
  element(selector).onclick({target: button, preventDefault() {}, stopPropagation() {}});
}
async function submit() {
  await element('#jf').onsubmit({preventDefault() {}});
  return submitted.getAll('files').map(file => file.name);
}

(async () => {
  const kind = process.argv[3];
  await viewJobForm(kind);
  if (kind === 'fill') element('#useSample').checked = true;
  const invoice = new File(['invoice'], '发票.xlsx', {lastModified: 1});
  const packing = new File(['packing'], '箱单.xlsx', {lastModified: 2});
  const contract = new File(['contract'], '合同.xlsx', {lastModified: 3});
  choose('#files', [invoice]);
  choose('#files', [packing]);
  assert.match(element('#dFiles').innerHTML, /发票.xlsx/, 'second selection must retain first file');
  drop('#dFiles', [contract]);
  choose('#files', [invoice]); // 重复选择同一个文件。
  choose('#files', []);        // 取消选择不能清空清单。
  assert.deepEqual(await submit(), ['发票.xlsx', '箱单.xlsx', '合同.xlsx']);
  assert.equal(element('#files').value, ''); // 原生选择框复位，允许移除后重选同一个文件。

  const before = element('#files').clicks;
  remove('#dFiles', 1);
  assert.equal(element('#files').clicks, before, 'remove must not reopen picker');
  assert.deepEqual(await submit(), ['发票.xlsx', '合同.xlsx']);
  choose('#files', [packing]);
  const revised = new File(['revised invoice'], '发票.xlsx', {lastModified: 4});
  const escaped = new File(['data'], '<img src=x>.xlsx', {lastModified: 5});
  choose('#files', [revised, escaped]);
  assert.deepEqual(await submit(), ['发票.xlsx', '合同.xlsx', '箱单.xlsx', '发票.xlsx', '<img src=x>.xlsx']);
  assert.match(element('#dFiles').innerHTML, /&lt;img/);
  assert.doesNotMatch(element('#dFiles').innerHTML, /<img/);
  assert.deepEqual(await Promise.all(submitted.getAll('files').map(file => file.text())),
    ['invoice', 'contract', 'packing', 'revised invoice', 'data']);
  for (let i = 0; i < 5; i++) remove('#dFiles', 0);
  assert.deepEqual(await submit(), []);
  assert.match(element('#dFiles').innerHTML, /点击选择/);
  drop('#dFiles', [packing, invoice]);
  assert.deepEqual(await submit(), ['箱单.xlsx', '发票.xlsx']);

  if (kind === 'fill') {
    const first = new File(['first'], '模板一.xls');
    const second = new File(['second'], '模板二.xlsx');
    choose('#tpl', [first]);
    drop('#dTpl', [second, first]);
    await submit();
    assert.equal(submitted.get('template').name, '模板二.xlsx');
    assert.equal(submitted.getAll('template').length, 1);
    remove('#dTpl', 0);
    await submit();
    assert.equal(submitted.get('template'), null);
    assert.match(element('#dTpl').innerHTML, /点击选择/);
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
