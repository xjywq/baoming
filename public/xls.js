// 最小的 .xls（BIFF8）读取器：只取单元格的值，满足小程序导出文件的需要
// 文件结构：OLE 复合文档（CFB）里的 Workbook 流 → BIFF8 记录
const CFB_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const CFB_HEADER_SIZE = 512;
const CFB_END_OF_CHAIN = 0xfffffffe;
const CFB_HEADER_DIFAT_COUNT = 109;
const CFB_DIR_ENTRY_SIZE = 128;
const CFB_DIR_NAME_BYTES = 64;
const CFB_ROOT_ENTRY_TYPE = 5;
const WORKBOOK_STREAM_NAME = 'Workbook';
const BIFF5_STREAM_NAME = 'Book';

const BIFF8_VERSION = 0x0600;
const RECORD = {
  BOF: 0x0809,
  EOF: 0x000a,
  BOUNDSHEET: 0x0085,
  SST: 0x00fc,
  CONTINUE: 0x003c,
  LABELSST: 0x00fd,
  LABEL: 0x0204,
  NUMBER: 0x0203,
  RK: 0x027e,
  MULRK: 0x00bd,
  FORMULA: 0x0006,
  BOOLERR: 0x0205,
  STRING: 0x0207,
};
const FORMULA_SPECIAL_MARK = 0xffff; // 结果不是数字时，FormulaValue 末两字节为 0xFFFF
const FORMULA_RESULT = { STRING: 0, BOOLEAN: 1, ERROR: 2, EMPTY: 3 };
const ERROR_TEXTS = { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x0f: '#VALUE!', 0x17: '#REF!', 0x1d: '#NAME?', 0x24: '#NUM!', 0x2a: '#N/A' };
const STRING_FLAG_HIGH_BYTE = 0x01;
const STRING_FLAG_EXT = 0x04;
const STRING_FLAG_RICH = 0x08;
const RICH_RUN_BYTES = 4;

function isXlsFile(bytes) {
  return CFB_SIGNATURE.every((byte, i) => bytes[i] === byte);
}

// ---------- CFB ----------

function readCfbStream(bytes, wantedName) {
  if (!isXlsFile(bytes)) throw new Error('不是 .xls 文件（文件头不对）');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const sectorSize = 1 << view.getUint16(0x1e, true);
  const miniSectorSize = 1 << view.getUint16(0x20, true);
  const fatSectorCount = view.getUint32(0x2c, true);
  const firstDirSector = view.getUint32(0x30, true);
  const miniStreamCutoff = view.getUint32(0x38, true);
  const firstMiniFatSector = view.getUint32(0x3c, true);
  let difatSector = view.getUint32(0x44, true);
  const sectorOffset = (sector) => CFB_HEADER_SIZE + sector * sectorSize;
  const maxSectors = Math.ceil(bytes.length / sectorSize);

  // FAT 扇区号：头部 109 个，超出部分在 DIFAT 链里
  const fatSectors = [];
  for (let i = 0; i < CFB_HEADER_DIFAT_COUNT && fatSectors.length < fatSectorCount; i++) {
    fatSectors.push(view.getUint32(0x4c + i * 4, true));
  }
  const entriesPerSector = sectorSize / 4;
  while (fatSectors.length < fatSectorCount && difatSector !== CFB_END_OF_CHAIN) {
    const base = sectorOffset(difatSector);
    for (let i = 0; i < entriesPerSector - 1 && fatSectors.length < fatSectorCount; i++) {
      fatSectors.push(view.getUint32(base + i * 4, true));
    }
    difatSector = view.getUint32(base + (entriesPerSector - 1) * 4, true);
  }
  const fat = [];
  for (const sector of fatSectors) {
    const base = sectorOffset(sector);
    for (let i = 0; i < entriesPerSector; i++) fat.push(view.getUint32(base + i * 4, true));
  }

  const followChain = (startSector, table, limit) => {
    const chain = [];
    for (let sector = startSector; sector !== CFB_END_OF_CHAIN; sector = table[sector]) {
      if (sector === undefined || chain.length > limit) throw new Error('.xls 文件已损坏（扇区链异常）');
      chain.push(sector);
    }
    return chain;
  };
  const readChain = (startSector) => {
    const chain = followChain(startSector, fat, maxSectors);
    const out = new Uint8Array(chain.length * sectorSize);
    chain.forEach((sector, i) => out.set(bytes.subarray(sectorOffset(sector), sectorOffset(sector) + sectorSize), i * sectorSize));
    return out;
  };

  const directory = readChain(firstDirSector);
  const dirView = new DataView(directory.buffer);
  const entries = [];
  for (let offset = 0; offset + CFB_DIR_ENTRY_SIZE <= directory.length; offset += CFB_DIR_ENTRY_SIZE) {
    const nameLength = dirView.getUint16(offset + CFB_DIR_NAME_BYTES, true);
    const nameChars = Math.max(0, nameLength / 2 - 1); // 长度含结尾的 \0
    let name = '';
    for (let i = 0; i < nameChars; i++) name += String.fromCharCode(dirView.getUint16(offset + i * 2, true));
    entries.push({
      name,
      type: directory[offset + 0x42],
      startSector: dirView.getUint32(offset + 0x74, true),
      size: dirView.getUint32(offset + 0x78, true),
    });
  }

  const entry = entries.find((e) => e.name === wantedName);
  if (!entry) {
    if (entries.some((e) => e.name === BIFF5_STREAM_NAME)) throw new Error('是 Excel 95 及更早的旧格式，不支持');
    throw new Error(`.xls 里找不到 ${wantedName} 数据流`);
  }
  if (entry.size >= miniStreamCutoff) return readChain(entry.startSector).subarray(0, entry.size);

  // 小于 4096 字节的流存放在 mini stream 里
  const root = entries.find((e) => e.type === CFB_ROOT_ENTRY_TYPE);
  const miniStream = readChain(root.startSector);
  const miniFatBytes = readChain(firstMiniFatSector);
  const miniFatView = new DataView(miniFatBytes.buffer);
  const miniFat = [];
  for (let i = 0; i < miniFatBytes.length / 4; i++) miniFat.push(miniFatView.getUint32(i * 4, true));
  const chain = followChain(entry.startSector, miniFat, miniFat.length);
  const out = new Uint8Array(chain.length * miniSectorSize);
  chain.forEach((sector, i) => out.set(miniStream.subarray(sector * miniSectorSize, (sector + 1) * miniSectorSize), i * miniSectorSize));
  return out.subarray(0, entry.size);
}

// ---------- BIFF8 ----------

function* iterateRecords(stream, start) {
  const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
  for (let offset = start; offset + 4 <= stream.length;) {
    const type = view.getUint16(offset, true);
    const length = view.getUint16(offset + 2, true);
    yield { type, offset, body: stream.subarray(offset + 4, offset + 4 + length) };
    offset += 4 + length;
  }
}

function decodeChars(bytes, offset, count, highByte) {
  let text = '';
  for (let i = 0; i < count; i++) {
    text += String.fromCharCode(highByte ? bytes[offset + i * 2] | (bytes[offset + i * 2 + 1] << 8) : bytes[offset + i]);
  }
  return text;
}

// SST 可能跨多个 CONTINUE 记录；字符被截断时，下一段开头会重新给出一个编码标志字节
function parseSharedStrings(segments) {
  let segmentIndex = 0;
  let position = 0;
  const current = () => segments[segmentIndex];
  const nextSegment = () => {
    segmentIndex++;
    position = 0;
    if (segmentIndex >= segments.length) throw new Error('.xls 共享字符串表不完整');
  };
  const readBytes = (count) => {
    const out = new Uint8Array(count);
    for (let filled = 0; filled < count;) {
      if (position >= current().length) nextSegment();
      const take = Math.min(count - filled, current().length - position);
      out.set(current().subarray(position, position + take), filled);
      filled += take;
      position += take;
    }
    return out;
  };
  const readUint16 = () => { const b = readBytes(2); return b[0] | (b[1] << 8); };
  const readUint32 = () => { const b = readBytes(4); return (b[0] | (b[1] << 8) | (b[2] << 16)) + b[3] * 0x1000000; };

  readUint32(); // 总引用次数
  const uniqueCount = readUint32();
  const strings = [];
  for (let n = 0; n < uniqueCount; n++) {
    const charCount = readUint16();
    const flags = readBytes(1)[0];
    let highByte = Boolean(flags & STRING_FLAG_HIGH_BYTE);
    const richRuns = flags & STRING_FLAG_RICH ? readUint16() : 0;
    const extBytes = flags & STRING_FLAG_EXT ? readUint32() : 0;
    let text = '';
    for (let remaining = charCount; remaining > 0;) {
      if (position >= current().length) {
        nextSegment();
        highByte = Boolean(current()[position++] & STRING_FLAG_HIGH_BYTE);
      }
      const bytesPerChar = highByte ? 2 : 1;
      const take = Math.min(remaining, Math.floor((current().length - position) / bytesPerChar));
      text += decodeChars(current(), position, take, highByte);
      position += take * bytesPerChar;
      remaining -= take;
    }
    readBytes(richRuns * RICH_RUN_BYTES + extBytes); // 富文本格式和扩展信息不需要
    strings.push(text);
  }
  return strings;
}

function decodeRk(rk) {
  const divideBy100 = rk & 0x01;
  const isInteger = rk & 0x02;
  let value;
  if (isInteger) {
    value = rk >> 2; // 有符号 30 位整数
  } else {
    const view = new DataView(new ArrayBuffer(8));
    view.setUint32(4, rk & 0xfffffffc, true); // RK 存的是 double 的高 30 位
    value = view.getFloat64(0, true);
  }
  return divideBy100 ? value / 100 : value;
}

function readXls(bytes) {
  const stream = readCfbStream(bytes, WORKBOOK_STREAM_NAME);
  const sheetEntries = [];
  let sharedStrings = [];

  // 工作簿全局部分：从开头到第一个 EOF
  const globals = [];
  for (const record of iterateRecords(stream, 0)) {
    globals.push(record);
    if (record.type === RECORD.EOF) break;
  }
  for (let i = 0; i < globals.length; i++) {
    const { type, body } = globals[i];
    const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
    if (type === RECORD.BOF && view.getUint16(0, true) !== BIFF8_VERSION) {
      throw new Error(`不支持的 Excel 版本（BIFF 0x${view.getUint16(0, true).toString(16)}）`);
    }
    if (type === RECORD.BOUNDSHEET) {
      const charCount = body[6];
      sheetEntries.push({ name: decodeChars(body, 8, charCount, body[7] & STRING_FLAG_HIGH_BYTE), offset: view.getUint32(0, true) });
    }
    if (type === RECORD.SST) {
      const segments = [body];
      while (globals[i + 1] && globals[i + 1].type === RECORD.CONTINUE) segments.push(globals[++i].body);
      sharedStrings = parseSharedStrings(segments);
    }
  }

  const sheets = {};
  for (const { name, offset } of sheetEntries) {
    sheets[name] = readSheetRows(stream, offset, sharedStrings, name);
  }
  return { sheetNames: sheetEntries.map((s) => s.name), sheets };
}

function errorText(code) {
  return ERROR_TEXTS[code] || `#ERR${code}`;
}

// 公式只取缓存的计算结果：普通数字直接存 double；字符串结果在紧跟的 STRING 记录里
function readFormulaResult(body, nextRecord, sheetName, row, col) {
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  if (view.getUint16(12, true) !== FORMULA_SPECIAL_MARK) return view.getFloat64(6, true);
  switch (body[6]) {
    case FORMULA_RESULT.STRING: {
      if (!nextRecord || nextRecord.type !== RECORD.STRING) {
        throw new Error(`工作表「${sheetName}」第 ${row + 1} 行第 ${col + 1} 列的公式缺少字符串结果`);
      }
      const charCount = nextRecord.body[0] | (nextRecord.body[1] << 8);
      return decodeChars(nextRecord.body, 3, charCount, nextRecord.body[2] & STRING_FLAG_HIGH_BYTE);
    }
    case FORMULA_RESULT.BOOLEAN:
      return Boolean(body[8]);
    case FORMULA_RESULT.ERROR:
      return errorText(body[8]);
    case FORMULA_RESULT.EMPTY:
      return '';
    default:
      throw new Error(`工作表「${sheetName}」第 ${row + 1} 行第 ${col + 1} 列的公式结果类型未知（${body[6]}）`);
  }
}

function readSheetRows(stream, start, sharedStrings, sheetName) {
  const rows = [];
  const put = (row, col, value) => {
    while (rows.length <= row) rows.push([]);
    rows[row][col] = value;
  };
  const records = [];
  for (const record of iterateRecords(stream, start)) {
    if (record.type === RECORD.EOF) break;
    records.push(record);
  }
  const cellTypes = [RECORD.LABELSST, RECORD.NUMBER, RECORD.RK, RECORD.MULRK, RECORD.LABEL, RECORD.FORMULA, RECORD.BOOLERR];
  for (let i = 0; i < records.length; i++) {
    const { type, body } = records[i];
    if (!cellTypes.includes(type)) continue;
    const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
    const row = view.getUint16(0, true);
    const col = view.getUint16(2, true);
    switch (type) {
      case RECORD.LABELSST:
        put(row, col, sharedStrings[view.getUint32(6, true)]);
        break;
      case RECORD.NUMBER:
        put(row, col, view.getFloat64(6, true));
        break;
      case RECORD.RK:
        put(row, col, decodeRk(view.getUint32(6, true)));
        break;
      case RECORD.MULRK: {
        const lastCol = view.getUint16(body.length - 2, true);
        for (let c = col; c <= lastCol; c++) put(row, c, decodeRk(view.getUint32(4 + (c - col) * 6 + 2, true)));
        break;
      }
      case RECORD.LABEL: {
        const charCount = view.getUint16(6, true);
        put(row, col, decodeChars(body, 9, charCount, body[8] & STRING_FLAG_HIGH_BYTE));
        break;
      }
      case RECORD.FORMULA:
        put(row, col, readFormulaResult(body, records[i + 1], sheetName, row, col));
        break;
      case RECORD.BOOLERR:
        put(row, col, body[7] ? errorText(body[6]) : Boolean(body[6]));
        break;
    }
  }
  return rows;
}

if (typeof module !== 'undefined') {
  module.exports = { readXls, isXlsFile, decodeRk };
}
