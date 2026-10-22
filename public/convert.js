// 预约名单转换：逻辑与原 Python 版 extract() 保持一致，浏览器和 node 共用
const SHEET_NAME = '文本数据';
const HEADER_ROW_INDEX = 1; // 第 1 行是活动标题，第 2 行是表头
const OUTPUT_COLUMNS = ['姓名', '性别', '手机号', '单位（部门）', '预约时间', '签到', '备注'];
const TIME_COLUMN = '预约时间';
const TIME_PATTERN = /(\d+):(\d{2})/;
const SELECTED_MARK = 1; // 时段列里填 1 表示预约了该时段
const MAX_HOUR = 23;
const MAX_MINUTE = 59;
const TITLE_FONT_SIZE = 16;
const BODY_FONT_SIZE = 12;
const COLUMN_WIDTHS = { 2: 25, 3: 15 }; // 列序号 → 字符宽度
const DAY_PATTERN = /(\d+)月(\d+)/;
const NUMERIC_TEXT_PATTERN = /^-?\d+(\.\d+)?$/;

function parseDay(text, year) {
  const match = text.match(DAY_PATTERN);
  if (!match) return null;
  return { year, month: Number(match[1]), day: Number(match[2]) };
}

// 先看文件名（与原脚本一致），链接下载的文件名可能没有日期，再看表格第一行的活动标题
function detectDay(filename, workbook, year) {
  const titleCell = getSheet(workbook)['A1'];
  return parseDay(filename, year) || parseDay(String(titleCell ? titleCell.v : ''), year);
}

function getSheet(workbook) {
  const sheet = workbook.Sheets[SHEET_NAME];
  if (!sheet) {
    throw new Error(`找不到工作表「${SHEET_NAME}」，文件里只有：${workbook.SheetNames.join('、')}`);
  }
  return sheet;
}

function formatIsoDate({ year, month, day }) {
  const pad = (n) => String(n).padStart(2, '0');
  return `${year}-${pad(month)}-${pad(day)}`;
}

function isEmptyCell(value) {
  return value === null || value === undefined || value === '';
}

// "9:00" → "09:00"；与 pd.to_datetime(format='%H:%M') 一样，非法时间直接报错
function normalizeTime(columnName) {
  const match = columnName.match(TIME_PATTERN);
  const hour = Number(match[1]);
  const minute = Number(match[2]);
  if (hour > MAX_HOUR || minute > MAX_MINUTE) {
    throw new Error(`时段列「${columnName}」里的时间 ${match[0]} 不合法`);
  }
  return `${String(hour).padStart(2, '0')}:${match[2]}`;
}

// 小程序导出的手机号是文本格式；pandas 读表时会把整列都是数字的文本转成数字，这里保持一致，
// 否则输出的手机号在 Excel 里会带“以文本形式存储的数字”警告
function convertNumericColumns(records) {
  for (const column of OUTPUT_COLUMNS) {
    if (column === TIME_COLUMN) continue;
    const values = records.map((record) => record[column]).filter((value) => !isEmptyCell(value));
    if (values.length === 0 || !values.every((value) => NUMERIC_TEXT_PATTERN.test(String(value)))) continue;
    for (const record of records) {
      if (!isEmptyCell(record[column])) record[column] = Number(record[column]);
    }
  }
}

function extractRecords(workbook) {
  const rows = XLSX.utils.sheet_to_json(getSheet(workbook), { header: 1, raw: true, defval: null });
  const header = rows[HEADER_ROW_INDEX];
  if (!header) {
    throw new Error(`工作表「${SHEET_NAME}」不足 ${HEADER_ROW_INDEX + 1} 行，找不到表头`);
  }

  const records = [];
  for (const row of rows.slice(HEADER_ROW_INDEX + 1)) {
    // 第一列是序号；空行和表尾的“由报名工具提供技术支持”都不是名单行
    if (!NUMERIC_TEXT_PATTERN.test(String(row[0] ?? ''))) continue;

    const record = {};
    for (const column of OUTPUT_COLUMNS) {
      const columnIndex = header.indexOf(column);
      if (columnIndex > 0) record[column] = row[columnIndex];
    }

    for (let columnIndex = 1; columnIndex < header.length; columnIndex++) {
      const columnName = header[columnIndex];
      if (OUTPUT_COLUMNS.includes(columnName)) continue;
      if (row[columnIndex] === SELECTED_MARK) {
        const nameText = String(columnName);
        record[TIME_COLUMN] = TIME_PATTERN.test(nameText) ? normalizeTime(nameText) : null;
        break;
      }
    }
    records.push(record);
  }
  convertNumericColumns(records);

  // 按预约时间升序，没有时间的排最后（与 pandas 的 NaT 行为一致）
  const timeOf = (record) => record[TIME_COLUMN];
  records.sort((a, b) => {
    if (isEmptyCell(timeOf(a))) return isEmptyCell(timeOf(b)) ? 0 : 1;
    if (isEmptyCell(timeOf(b))) return -1;
    return timeOf(a).localeCompare(timeOf(b));
  });
  return records;
}

function buildOutputWorkbook(records, day) {
  const thinBorder = { style: 'thin', color: { rgb: '000000' } };
  const border = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
  const titleStyle = {
    font: { bold: true, sz: TITLE_FONT_SIZE },
    alignment: { horizontal: 'center', vertical: 'center' },
    border,
  };
  const headerStyle = {
    font: { bold: true, sz: BODY_FONT_SIZE },
    alignment: { horizontal: 'center' },
    border,
  };
  const bodyStyle = {
    font: { sz: BODY_FONT_SIZE },
    alignment: { horizontal: 'center' },
    border,
  };

  const columnCount = OUTPUT_COLUMNS.length;
  const title = `${formatIsoDate(day)}预约名单`;
  const titleRow = [title, ...Array(columnCount - 1).fill('')];
  const bodyRows = records.map((record) =>
    OUTPUT_COLUMNS.map((column) => (isEmptyCell(record[column]) ? '' : record[column])),
  );
  const sheet = XLSX.utils.aoa_to_sheet([titleRow, OUTPUT_COLUMNS, ...bodyRows]);

  const rowStyles = [titleStyle, headerStyle];
  const totalRows = 2 + bodyRows.length;
  for (let rowIndex = 0; rowIndex < totalRows; rowIndex++) {
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      sheet[address].s = rowStyles[rowIndex] || bodyStyle;
    }
  }
  sheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: columnCount - 1 } }];
  sheet['!cols'] = OUTPUT_COLUMNS.map((_, columnIndex) =>
    COLUMN_WIDTHS[columnIndex] ? { wch: COLUMN_WIDTHS[columnIndex] } : {},
  );

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, SHEET_NAME);
  return workbook;
}

function outputFilename(day) {
  return `${day.month}月${day.day}日.xlsx`;
}

if (typeof module !== 'undefined') {
  module.exports = { detectDay, formatIsoDate, extractRecords, buildOutputWorkbook, outputFilename, OUTPUT_COLUMNS };
}
