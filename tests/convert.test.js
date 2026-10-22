// 运行：node --test
const test = require('node:test');
const assert = require('node:assert');
const XLSX = require('xlsx-js-style'); // 仅测试用：解析生成的 .xlsx 做对照
const { detectDay, extractRecords, buildOutputSheet, outputFilename, OUTPUT_COLUMNS, OUTPUT_FONTS } = require('../public/convert.js');
const { writeXlsx } = require('../public/xlsx.js');

const TITLE = '10月22日上海航天149厂浦江园区出诊活动 (2026-10-22 08:33:25)';
const SAMPLE_HEADER = ['序号', '昵称', '姓名', '性别', '手机号', '单位（部门）', '9:00-9:30', '13:30-14:00', '10:00-10:30'];

// 与 readXls 返回值同构
function sampleWorkbook(dataRows, title = TITLE) {
  return { sheetNames: ['文本数据'], sheets: { 文本数据: [[title], SAMPLE_HEADER, ...dataRows] } };
}

test('按时段提取并排序，未选时段的排最后，跳过空行和表尾说明', () => {
  const records = extractRecords(sampleWorkbook([
    ['1', '小A', '张三', '女', '13800000001', '一分厂', undefined, 1],
    ['2', '小B', '李四', '男', '13800000002', '一分厂', 1],
    ['3', '1', '王五', '男', '13800000003', '一分厂'],
    ['4', '小D', '赵六', '男', '13800000004', '一分厂', undefined, undefined, 1],
    [],
    ['由报名工具提供技术支持'],
  ]));
  assert.deepStrictEqual(records.map((r) => [r['姓名'], r['预约时间']]), [
    ['李四', '09:00'],
    ['赵六', '10:00'],
    ['张三', '13:30'],
    ['王五', undefined],
  ]);
  assert.strictEqual(records[0]['手机号'], 13800000002, '整列都是数字的文本转成数字，与原脚本 pandas 行为一致');
});

test('输出 xlsx：标题合并、表头、数据行、样式、列宽', () => {
  const records = extractRecords(sampleWorkbook([['1', '小A', '张三', '女', '13800000001', '一分厂', 1]]));
  const day = { year: 2026, month: 10, day: 22 };
  const bytes = writeXlsx(buildOutputSheet(records, day), OUTPUT_FONTS);
  const workbook = XLSX.read(bytes, { type: 'array', cellStyles: true });
  assert.deepStrictEqual(workbook.SheetNames, ['文本数据']);
  const sheet = workbook.Sheets['文本数据'];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  assert.strictEqual(rows[0][0], '2026-10-22预约名单');
  assert.deepStrictEqual(rows[1], OUTPUT_COLUMNS);
  assert.deepStrictEqual(rows[2], ['张三', '女', 13800000001, '一分厂', '09:00', '', '']);
  assert.strictEqual(rows.length, 3);
  assert.deepStrictEqual(sheet['!merges'], [{ s: { r: 0, c: 0 }, e: { r: 0, c: 6 } }]);
  assert.strictEqual(sheet['!cols'][2].wch, 25);
  assert.strictEqual(sheet['!cols'][3].wch, 15);
  assert.strictEqual(outputFilename(day), '10月22日.xlsx');
});

test('缺少「文本数据」工作表时报错并列出现有工作表', () => {
  const workbook = { sheetNames: ['Sheet1'], sheets: { Sheet1: [['x']] } };
  assert.throws(() => extractRecords(workbook), /找不到工作表「文本数据」.*Sheet1/);
});

test('日期优先取文件名，没有再取活动标题', () => {
  const workbook = sampleWorkbook([]);
  assert.deepStrictEqual(detectDay('11月3日.xls', workbook, 2026), { year: 2026, month: 11, day: 3 });
  assert.deepStrictEqual(detectDay('导出链接', workbook, 2026), { year: 2026, month: 10, day: 22 });
  assert.strictEqual(detectDay('导出链接', sampleWorkbook([], '出诊活动'), 2026), null);
});
