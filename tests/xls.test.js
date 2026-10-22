const test = require('node:test');
const assert = require('node:assert');
const XLSX = require('xlsx-js-style'); // 仅测试用：生成 .xls、封装 OLE 复合文档
const { readXls, isXlsFile, decodeRk } = require('../public/xls.js');

// ---------- 手工构造 BIFF8 记录 ----------
const u16 = (n) => [n & 0xff, (n >> 8) & 0xff];
const u32 = (n) => [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >>> 24) & 0xff];
const f64 = (x) => { const b = Buffer.alloc(8); b.writeDoubleLE(x); return [...b]; };
const utf16 = (text) => [...Buffer.from(text, 'utf16le')];
const latin1 = (text) => [...Buffer.from(text, 'latin1')];
const record = (type, body) => [...u16(type), ...u16(body.length), ...body];
const bof = (dt) => record(0x0809, [...u16(0x0600), ...u16(dt), ...u16(0), ...u16(1997), ...u32(0), ...u32(0x0600)]);
const eof = () => record(0x000a, []);
const cell = (row, col) => [...u16(row), ...u16(col), ...u16(0)];
const rkInt = (n, times100 = false) => ((n << 2) | 0x02 | (times100 ? 1 : 0)) >>> 0;
const formulaSpecial = (row, col, kind, value = 0) =>
  record(0x0006, [...cell(row, col), kind, 0, value, 0, 0, 0, 0xff, 0xff, ...u16(0), ...u32(0), ...u16(0)]);

// 共享字符串表故意拆成多段 CONTINUE：
//   'ABCDEF' 前半段单字节，CONTINUE 开头的标志字节切换成双字节；
//   '中文名字' 是富文本（带 1 个格式段），字符跨段；'x' 的字符串头正好在新段开头
function sharedStringRecords() {
  const sst = record(0x00fc, [...u32(3), ...u32(3), ...u16(6), 0x00, ...latin1('ABC')]);
  const continue1 = record(0x003c, [0x01, ...utf16('DEF'), ...u16(4), 0x09, ...u16(1), ...utf16('中文')]);
  const continue2 = record(0x003c, [0x01, ...utf16('名字'), 0, 0, 0, 0]);
  const continue3 = record(0x003c, [...u16(1), 0x00, ...latin1('x')]);
  return [...sst, ...continue1, ...continue2, ...continue3];
}

function sheetRecords(extraRows) {
  const records = [
    ...bof(0x0010),
    ...record(0x00fd, [...cell(0, 0), ...u32(0)]),
    ...record(0x00fd, [...cell(1, 0), ...u32(1)]),
    ...record(0x027e, [...cell(1, 1), ...u32(rkInt(42))]),
    ...record(0x027e, [...cell(1, 2), ...u32(rkInt(1234, true))]),
    ...record(0x00bd, [...u16(2), ...u16(0), ...u16(0), ...u32(rkInt(-7)), ...u16(0), ...u32(0x3ff80000), ...u16(1)]),
    ...record(0x0006, [...cell(3, 0), ...f64(3.25), ...u16(0), ...u32(0), ...u16(0)]),
    ...formulaSpecial(3, 1, 0),
    ...record(0x0207, [...u16(2), 0x00, ...latin1('ok')]),
    ...formulaSpecial(3, 2, 3),
    ...record(0x0205, [...cell(3, 3), 1, 0]),
    ...record(0x0204, [...cell(3, 4), ...u16(2), 0x01, ...utf16('标签')]),
    ...record(0x00fd, [...cell(4, 0), ...u32(2)]),
  ];
  for (let i = 0; i < extraRows; i++) records.push(...record(0x0203, [...cell(5 + i, 0), ...f64(-2.5)]));
  return [...records, ...eof()];
}

function buildXls(extraRows = 0) {
  const sheetName = utf16('文本数据');
  const boundsheetLength = 4 + 8 + sheetName.length;
  const beforeBoundsheet = bof(0x0005);
  const afterBoundsheet = [...sharedStringRecords(), ...eof()];
  const sheetOffset = beforeBoundsheet.length + boundsheetLength + afterBoundsheet.length;
  const boundsheet = record(0x0085, [...u32(sheetOffset), 0, 0, 4, 0x01, ...sheetName]);
  const stream = Buffer.from([...beforeBoundsheet, ...boundsheet, ...afterBoundsheet, ...sheetRecords(extraRows)]);

  const container = XLSX.CFB.utils.cfb_new();
  XLSX.CFB.utils.cfb_add(container, 'Workbook', stream);
  return { bytes: new Uint8Array(XLSX.CFB.write(container, { type: 'buffer' })), streamLength: stream.length };
}

const EXPECTED_ROWS = [
  ['ABCDEF'],
  ['中文名字', 42, 12.34],
  [-7, 1.5],
  [3.25, 'ok', '', true, '标签'],
  ['x'],
];

test('读取手工构造的 BIFF8：SST 跨 CONTINUE、RK、MULRK、公式结果、布尔、LABEL', () => {
  const { bytes, streamLength } = buildXls();
  assert.ok(streamLength < 4096, '小于 4096 字节，走 mini stream');
  const workbook = readXls(bytes);
  assert.deepStrictEqual(workbook.sheetNames, ['文本数据']);
  assert.deepStrictEqual(workbook.sheets['文本数据'].map((row) => Array.from(row)), EXPECTED_ROWS);
});

test('Workbook 流超过 4096 字节时走普通 FAT', () => {
  const extraRows = 400;
  const { bytes, streamLength } = buildXls(extraRows);
  assert.ok(streamLength >= 4096);
  const rows = readXls(bytes).sheets['文本数据'];
  assert.deepStrictEqual(rows.slice(0, 5).map((row) => Array.from(row)), EXPECTED_ROWS);
  assert.strictEqual(rows.length, 5 + extraRows);
  assert.deepStrictEqual(rows[rows.length - 1], [-2.5]);
});

test('读取 SheetJS 写出的大文件（LABEL / NUMBER 记录，多个 FAT 扇区）', () => {
  const sourceRows = [['标题'], ['序号', '姓名']];
  for (let i = 0; i < 3000; i++) sourceRows.push([i + 1, `测试姓名${i}`]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(sourceRows), '文本数据');
  const bytes = new Uint8Array(XLSX.write(workbook, { type: 'array', bookType: 'biff8' }));
  assert.deepStrictEqual(readXls(bytes).sheets['文本数据'].map((row) => Array.from(row)), sourceRows);
});

test('RK 编码', () => {
  assert.strictEqual(decodeRk(rkInt(-123)), -123);
  assert.strictEqual(decodeRk(rkInt(150, true)), 1.5);
  assert.strictEqual(decodeRk(0x3ff80000), 1.5);
});

test('不是 .xls 时报错', () => {
  const html = new TextEncoder().encode('<!DOCTYPE html><html></html>');
  assert.ok(!isXlsFile(html));
  assert.throws(() => readXls(html), /不是 \.xls 文件/);
});
