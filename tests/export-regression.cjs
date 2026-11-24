'use strict';

const { readFileSync } = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const { test } = require('node:test');
const XLSX = require('../xlsx.full.min.js');

const html = readFileSync(require.resolve('../index.html'), 'utf8');
const [, excelJS, app] = [...html.matchAll(/<script[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
function section(start, end) {
  const i = app.indexOf(start);
  const j = app.indexOf(end, i);
  assert.ok(i >= 0 && j > i, `找不到待测试代码段：${start}`);
  return app.slice(i, j);
}

function workbook(rows, configure = () => {}, compression = false) {
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  configure(sheet);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheet, '数据');
  return XLSX.write(wb, { type: 'array', bookType: 'xlsx', compression });
}
function read(buf) { return XLSX.read(buf, { type: 'array', cellFormula: true, cellDates: false }).Sheets.数据; }

// 运行页面的实际导出函数和内嵌 ExcelJS；用内存替代浏览器存储与本地文件句柄。
function fixture(bytes, annotations = []) {
  const state = {
    fileName: '测试.xlsx', fileHash: 'fixture', fileBytes: bytes,
    wb: XLSX.read(bytes, { type: 'array' }), sheetNames: ['数据'],
    sheetOrigins: { 数据: { r: 0, c: 0 } }, sheetState: { 数据: { edits: new Map() } },
  };
  let signal;
  const ctx = vm.createContext({
    XLSX, state, console, Blob, TextEncoder, TextDecoder, ArrayBuffer, Uint8Array,
    setTimeout, clearTimeout, setImmediate, IDBKeyRange: { bound() {} },
    annIdbOp: (_, fn) => fn({
      getAllKeys: () => annotations.length ? ['测试.xlsx::fixture::数据'] : [],
      getAll: () => annotations.length ? [{ val: annotations }] : [],
    }),
    saveAnns: async () => {}, countAllAnns: async () => annotations.length, commitEditing() {},
    toast: (text, kind) => signal({ text, kind }),
  });
  ctx.window = ctx;
  ctx.self = ctx;
  vm.runInContext(excelJS, ctx);
  vm.runInContext([
    section('function hashBytes(', 'const enc ='),
    section('const ANN_COLORS =', 'let annCache'),
    section('const THEME_MAX_UNCOMPRESSED', 'function loadFile('),
    section('function exportStats()', "$('#btnExport').onclick"),
    section('async function buildExportBuffer(', "$('#btnReturnSource').onclick"),
  ].join('\n'), ctx);
  return {
    edits: state.sheetState.数据.edits,
    export: () => vm.runInContext('buildExportBuffer()', ctx),
    inspect: bytes => vm.runInContext('(async bytes => { const w = new ExcelJS.Workbook(); await w.xlsx.load(bytes); const s = w.getWorksheet("数据"); return Object.fromEntries(["A2", "B2", "C2"].map(a => { const c = s.getCell(a); return [a, { formula: c.value?.formula, value: c.value?.result ?? c.value, color: c.font?.color?.argb, note: c.note, richText: c.value?.richText }]; })); })', ctx)(bytes),
    async returnToSource(candidate, changedDuringExport = null) {
      let written;
      const file = (buf, lastModified) => new File([buf], state.fileName, { lastModified });
      ctx.target = file(candidate, 1);
      ctx.handle = {
        async getFile() { return changedDuringExport ? file(changedDuringExport, 1) : ctx.target; },
        async createWritable() { return { async write(buf) { written = buf; }, async close() {} }; },
      };
      const toast = new Promise(resolve => { signal = resolve; });
      await vm.runInContext('returnToSourceFile(target, handle)', ctx);
      return { ...(await toast), written };
    },
  };
}

const rows = [['编号', '备注'], ['ID-A', '原始A'], ['ID-B', '原始B']];

test('回传拒绝同名同工作表、但已经排序或改动的源文件', async () => {
  const app = fixture(workbook(rows));
  app.edits.set('1,1', '只修改 ID-A');
  for (const candidate of [
    workbook([rows[0], rows[2], rows[1]]),
    workbook([rows[0], rows[1], ['ID-B', '外部修改']]),
  ]) {
    const result = await app.returnToSource(candidate);
    assert.equal(!!result.written, false);
    assert.equal(result.kind, 'err');
  }
});

test('回传允许内容未变但 ZIP 字节不同的副本；同一会话可回写撤销', async () => {
  const source = workbook(rows);
  const equivalent = workbook(rows, undefined, true);
  assert.notEqual(source.byteLength, equivalent.byteLength);
  const app = fixture(source);
  app.edits.set('1,1', '只修改 ID-A');
  const result = await app.returnToSource(equivalent);
  assert.equal(read(result.written).B2.v, '只修改 ID-A');
  assert.equal(read(result.written).B3.v, '原始B');
  app.edits.clear();
  const reverted = await app.returnToSource(result.written);
  assert.equal(read(reverted.written).B2.v, '原始A');
});

test('回传生成期间源文件被另一个程序修改时不覆盖', async () => {
  const source = workbook(rows);
  const app = fixture(source);
  app.edits.set('1,1', '只修改 ID-A');
  const changed = workbook([rows[0], rows[2], rows[1]]);
  assert.equal(source.byteLength, changed.byteLength);
  const result = await app.returnToSource(source, changed);
  assert.equal(!!result.written, false);
  assert.equal(result.kind, 'err');
});

test('标注回传只写颜色、不添加批注；公式数字保留并整格着色', async () => {
  const source = workbook([['数字', '公式', '文本'], [3, 6, '标注文本']], s => { s.B2.f = 'A2*2'; });
  const app = fixture(source, [
    ['1,0', [{ s: 0, e: 1, ci: 0, anchor: '3' }]],
    ['1,1', [{ s: 0, e: 1, ci: 0, anchor: '6' }]],
    ['1,2', [{ s: 0, e: 2, ci: 0, anchor: '标注' }]],
  ]);
  const bytes = await app.export();
  const result = read(bytes);
  const styles = await app.inspect(bytes);
  assert.equal(result.B2.f, 'A2*2');
  assert.equal(result.B2.v, 6);
  assert.equal(result.B2.c, undefined);
  assert.equal(styles.B2.color, 'FFD54941');
  assert.equal(result.A2.t, 'n');
  assert.equal(result.A2.v, 3);
  assert.equal(result.A2.c, undefined);
  assert.equal(styles.A2.color, 'FFD54941');
  assert.equal(result.C2.v, '标注文本');
  assert.match(result.C2.r, /<rPr>/);
  assert.equal(result.C2.c, undefined);
  assert.equal(styles.C2.note, undefined);
});

test('只有公式被标注时也保留公式并着色，不添加批注', async () => {
  const source = workbook([['数字', '公式'], [3, 6]], s => { s.B2.f = 'A2*2'; });
  const app = fixture(source, [['1,1', [{ s: 0, e: 1, ci: 0, anchor: '6' }]]]);
  const bytes = await app.export();
  const result = read(bytes);
  const styles = await app.inspect(bytes);
  assert.equal(result.B2.f, 'A2*2');
  assert.equal(result.B2.v, 6);
  assert.equal(result.B2.c, undefined);
  assert.equal(styles.B2.color, 'FFD54941');
});
