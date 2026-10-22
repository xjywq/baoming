// 最小的 .xlsx 生成器：单个工作表、内联字符串、少量固定样式，zip 用不压缩的 STORE 方式
const XLSX_STYLE = { DEFAULT: 0, TITLE: 1, HEADER: 2, BODY: 3 };
const ZIP_LOCAL_HEADER_SIG = 0x04034b50;
const ZIP_CENTRAL_HEADER_SIG = 0x02014b50;
const ZIP_END_SIG = 0x06054b50;
const ZIP_VERSION = 20;
const ZIP_UTF8_FLAG = 0x0800;
const DOS_EPOCH_YEAR = 1980;
const COLUMN_WIDTH_PADDING = 0.7109375; // Excel 列宽含单元格内边距，与原脚本 xlsxwriter 的写法一致
const INVALID_XML_CHARS = /[\x00-\x08\x0B\x0C\x0E-\x1F]/g;

function escapeXml(value) {
  return String(value)
    .replace(INVALID_XML_CHARS, '')
    .replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[ch]);
}

function columnLetter(index) {
  let letters = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  return letters;
}

function cellXml(rowIndex, columnIndex, value, style) {
  const ref = `${columnLetter(columnIndex)}${rowIndex + 1}`;
  if (value === null || value === undefined || value === '') return `<c r="${ref}" s="${style}"/>`;
  if (typeof value === 'number') return `<c r="${ref}" s="${style}"><v>${value}</v></c>`;
  return `<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
}

// sheet: { name, rows: [[值]], rowStyles: [样式号], merges: ['A1:G1'], columnWidths: { 列序号: 字符数 } }
function sheetXml({ rows, rowStyles, merges, columnWidths }) {
  const cols = Object.entries(columnWidths)
    .map(([index, chars]) => `<col min="${Number(index) + 1}" max="${Number(index) + 1}" width="${chars + COLUMN_WIDTH_PADDING}" customWidth="1"/>`)
    .join('');
  const rowsXml = rows
    .map((row, r) => `<row r="${r + 1}">${row.map((value, c) => cellXml(r, c, value, rowStyles[r])).join('')}</row>`)
    .join('');
  const mergesXml = merges.length ? `<mergeCells count="${merges.length}">${merges.map((ref) => `<mergeCell ref="${ref}"/>`).join('')}</mergeCells>` : '';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + (cols ? `<cols>${cols}</cols>` : '')
    + `<sheetData>${rowsXml}</sheetData>${mergesXml}</worksheet>`;
}

function stylesXml(titleFontSize, bodyFontSize) {
  const thin = '<left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/>';
  const xf = (fontId, borderId, alignment) => `<xf numFmtId="0" fontId="${fontId}" fillId="0" borderId="${borderId}" xfId="0"`
    + (borderId ? ' applyBorder="1"' : '') + (fontId ? ' applyFont="1"' : '')
    + (alignment ? ` applyAlignment="1"><alignment ${alignment}/></xf>` : '/>');
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + '<fonts count="4">'
    + '<font><sz val="11"/><name val="Calibri"/></font>'
    + `<font><b/><sz val="${titleFontSize}"/><name val="Calibri"/></font>`
    + `<font><b/><sz val="${bodyFontSize}"/><name val="Calibri"/></font>`
    + `<font><sz val="${bodyFontSize}"/><name val="Calibri"/></font>`
    + '</fonts>'
    + '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
    + `<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border>${thin}</border></borders>`
    + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    + '<cellXfs count="4">'
    + xf(0, 0, '')
    + xf(1, 1, 'horizontal="center" vertical="center"')
    + xf(2, 1, 'horizontal="center"')
    + xf(3, 1, 'horizontal="center"')
    + '</cellXfs>'
    + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    + '</styleSheet>';
}

function workbookFiles(sheet, fonts) {
  const xmlHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  return {
    '[Content_Types].xml': `${xmlHeader}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
      + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
      + '</Types>',
    '_rels/.rels': `${xmlHeader}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
      + `<Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `${xmlHeader}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${rel}">`
      + `<sheets><sheet name="${escapeXml(sheet.name)}" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `${xmlHeader}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
      + `<Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/>`
      + `<Relationship Id="rId2" Type="${rel}/styles" Target="styles.xml"/></Relationships>`,
    'xl/styles.xml': stylesXml(fonts.title, fonts.body),
    'xl/worksheets/sheet1.xml': sheetXml(sheet),
  };
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function dosDateTime(date) {
  const time = (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2);
  const day = ((date.getFullYear() - DOS_EPOCH_YEAR) << 9) | ((date.getMonth() + 1) << 5) | date.getDate();
  return { time, day };
}

function zipStore(files, date) {
  const encoder = new TextEncoder();
  const { time, day } = dosDateTime(date);
  const entries = Object.entries(files).map(([name, content]) => {
    const data = encoder.encode(content);
    return { name: encoder.encode(name), data, crc: crc32(data) };
  });
  const localSize = entries.reduce((sum, e) => sum + 30 + e.name.length + e.data.length, 0);
  const centralSize = entries.reduce((sum, e) => sum + 46 + e.name.length, 0);
  const out = new Uint8Array(localSize + centralSize + 22);
  const view = new DataView(out.buffer);
  let offset = 0;
  const writeCommon = (entry) => { // 本地头和中央目录共用的字段：版本、标志、方法、时间、CRC、大小、文件名长度
    view.setUint16(offset, ZIP_VERSION, true);
    view.setUint16(offset + 2, ZIP_UTF8_FLAG, true);
    view.setUint16(offset + 4, 0, true); // STORE
    view.setUint16(offset + 6, time, true);
    view.setUint16(offset + 8, day, true);
    view.setUint32(offset + 10, entry.crc, true);
    view.setUint32(offset + 14, entry.data.length, true);
    view.setUint32(offset + 18, entry.data.length, true);
    view.setUint16(offset + 22, entry.name.length, true);
  };
  for (const entry of entries) {
    entry.offset = offset;
    view.setUint32(offset, ZIP_LOCAL_HEADER_SIG, true);
    offset += 4;
    writeCommon(entry);
    offset += 26;
    out.set(entry.name, offset);
    offset += entry.name.length;
    out.set(entry.data, offset);
    offset += entry.data.length;
  }
  const centralStart = offset;
  for (const entry of entries) {
    view.setUint32(offset, ZIP_CENTRAL_HEADER_SIG, true);
    view.setUint16(offset + 4, ZIP_VERSION, true);
    offset += 6;
    writeCommon(entry);
    view.setUint32(offset + 36, entry.offset, true); // 其余字段（注释、属性等）保持 0
    offset += 40;
    out.set(entry.name, offset);
    offset += entry.name.length;
  }
  view.setUint32(offset, ZIP_END_SIG, true);
  view.setUint16(offset + 8, entries.length, true);
  view.setUint16(offset + 10, entries.length, true);
  view.setUint32(offset + 12, offset - centralStart, true);
  view.setUint32(offset + 16, centralStart, true);
  return out;
}

function writeXlsx(sheet, fonts, date = new Date()) {
  return zipStore(workbookFiles(sheet, fonts), date);
}

if (typeof module !== 'undefined') {
  module.exports = { writeXlsx, XLSX_STYLE, columnLetter };
}
