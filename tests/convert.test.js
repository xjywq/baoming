// 运行：node --test tests/convert.test.js
const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

// 和浏览器一样按普通脚本加载（定义全局 XLSX）；按 CommonJS 加载会去 require 已删除的 cpexcel.js
vm.runInThisContext(fs.readFileSync(path.join(__dirname, '../public/xlsx.bundle.js'), 'utf8'));
const { detectDay, extractRecords, buildOutputWorkbook, outputFilename, OUTPUT_COLUMNS } = require('../public/convert.js');

const SAMPLE_HEADER = ['序号', '昵称', '姓名', '性别', '手机号', '单位（部门）', '9:00-9:30', '13:30-14:00', '10:00-10:30'];

// 构造一个和小程序导出结构一致的 .xls（BIFF8），再读回来，覆盖真实的读取路径
function sampleWorkbook(dataRows) {
  const title = '10月22日上海航天149厂浦江园区出诊活动 (2026-10-22 08:33:25)';
  const sheet = XLSX.utils.aoa_to_sheet([[title], SAMPLE_HEADER, ...dataRows]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, '文本数据');
  const bytes = XLSX.write(workbook, { type: 'array', bookType: 'biff8' });
  return XLSX.read(bytes, { type: 'array' });
}

test('按时段提取并排序，未选时段的排最后', () => {
  const workbook = sampleWorkbook([
    [1, '小A', '张三', '女', '13800000001', '一分厂', null, 1, null],
    [2, '小B', '李四', '男', '13800000002', '一分厂', 1, null, null],
    [3, '1', '王五', '男', '13800000003', '一分厂', null, null, null],
    [4, '小D', '赵六', '男', '13800000004', '一分厂', null, null, 1],
    [null, null, null, null, null, null, null, null, null],
    ['由报名工具提供技术支持', null, null, null, null, null, null, null, null],
  ]);
  const records = extractRecords(workbook);
  assert.deepStrictEqual(records.map((r) => [r['姓名'], r['预约时间']]), [
    ['李四', '09:00'],
    ['赵六', '10:00'],
    ['张三', '13:30'],
    ['王五', undefined],
  ]);
  assert.strictEqual(records[0]['手机号'], 13800000002);
});

test('输出表格：标题合并、表头、数据行', () => {
  const records = extractRecords(sampleWorkbook([[1, '小A', '张三', '女', '13800000001', '一分厂', 1, null, null]]));
  const day = { year: 2026, month: 10, day: 22 };
  const bytes = XLSX.write(buildOutputWorkbook(records, day), { type: 'array', bookType: 'xlsx' });
  const sheet = XLSX.read(bytes, { type: 'array' }).Sheets['文本数据'];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  assert.strictEqual(rows[0][0], '2026-10-22预约名单');
  assert.deepStrictEqual(rows[1], OUTPUT_COLUMNS);
  assert.deepStrictEqual(rows[2], ['张三', '女', 13800000001, '一分厂', '09:00', '', '']);
  assert.strictEqual(rows.length, 3);
  assert.deepStrictEqual(sheet['!merges'], [{ s: { r: 0, c: 0 }, e: { r: 0, c: 6 } }]);
  assert.strictEqual(outputFilename(day), '10月22日.xlsx');
});

test('缺少「文本数据」工作表时报错并列出现有工作表', () => {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['x']]), 'Sheet1');
  assert.throws(() => extractRecords(workbook), /找不到工作表「文本数据」.*Sheet1/);
});

test('日期优先取文件名，没有再取活动标题', () => {
  const workbook = sampleWorkbook([]);
  assert.deepStrictEqual(detectDay('11月3日.xls', workbook, 2026), { year: 2026, month: 11, day: 3 });
  assert.deepStrictEqual(detectDay('download.xls', workbook, 2026), { year: 2026, month: 10, day: 22 });

  const untitled = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(untitled, XLSX.utils.aoa_to_sheet([['出诊活动'], SAMPLE_HEADER]), '文本数据');
  assert.strictEqual(detectDay('download.xls', untitled, 2026), null);
});
