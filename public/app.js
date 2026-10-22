const PICKER_ID = 'roster-source'; // 浏览器按此 id 记住上次选文件的目录
const SPREADSHEET_TYPES = [{
  description: '报名表',
  accept: {
    'application/vnd.ms-excel': ['.xls'],
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
  },
}];
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const LINK_SOURCE_NAME = '导出链接'; // 接口不允许网页读取文件名，日期从表格标题识别
const BLOB_URL_LIFETIME_MS = 60000;
const NO_SLOT_LABEL = '未选时段';

const $ = (id) => document.getElementById(id);
const fileInput = $('file-input');
const linkForm = $('link-form');
const linkInput = $('link-input');
const dateInput = $('date-input');
const statusBox = $('status');

let currentRecords = null;
// iOS Safari 要求点击后立刻调用 navigator.share，否则报 NotAllowedError，所以输出文件提前生成好
let preparedFile = null;
let prepareError = null;

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[ch]);
}

function showStatus(message, kind) {
  statusBox.textContent = message;
  statusBox.className = `status ${kind}`;
  statusBox.hidden = false;
}

function clearStatus() {
  statusBox.hidden = true;
}

function selectedDay() {
  if (!dateInput.value) throw new Error('请先选择活动日期');
  const [year, month, day] = dateInput.value.split('-').map(Number);
  return { year, month, day };
}

function loadWorkbook(bytes, sourceName) {
  const workbook = XLSX.read(bytes, { type: 'array' });
  currentRecords = extractRecords(workbook);
  const day = detectDay(sourceName, workbook, new Date().getFullYear());
  dateInput.value = day ? formatIsoDate(day) : '';
  prepareOutput();
  renderResult(sourceName, Boolean(day));
}

function prepareOutput() {
  try {
    preparedFile = buildOutputFile();
    prepareError = null;
  } catch (error) {
    preparedFile = null;
    prepareError = error;
    console.error(error);
  }
}

// 没有可用的输出文件时提示原因，返回 null
function outputFileOrReport() {
  if (!preparedFile) showStatus(`生成失败：${prepareError.message}`, 'error');
  return preparedFile;
}

function renderResult(sourceName, dayDetected) {
  $('source-name').textContent = sourceName;
  $('record-count').textContent = currentRecords.length;

  const slotCounts = new Map();
  for (const record of currentRecords) {
    const slot = record['预约时间'] || NO_SLOT_LABEL;
    slotCounts.set(slot, (slotCounts.get(slot) || 0) + 1);
  }
  $('slot-chips').innerHTML = [...slotCounts]
    .map(([slot, count]) => `<span class="chip">${escapeHtml(slot)}<b>${count}人</b></span>`)
    .join('');

  const rowsHtml = currentRecords.map((record) => `<tr>
      <td class="time">${escapeHtml(record['预约时间'])}</td>
      <td>${escapeHtml(record['姓名'])}</td>
      <td>${escapeHtml(record['单位（部门）'])}</td>
      <td>${escapeHtml(record['手机号'])}</td>
    </tr>`).join('');
  $('preview').innerHTML = `<tr><th>时间</th><th>姓名</th><th>单位</th><th>手机号</th></tr>${rowsHtml}`;

  $('result').hidden = false;
  $('action-bar').hidden = false;
  if (dayDetected) clearStatus();
  else showStatus('文件名和标题里都没找到日期，请手动选择活动日期', 'error');
  $('result').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

async function handleFile(file) {
  clearStatus();
  try {
    loadWorkbook(await file.arrayBuffer(), file.name);
  } catch (error) {
    showStatus(`读取「${file.name}」失败：${error.message}`, 'error');
    console.error(error);
  }
}

// Chrome / Edge 用系统文件选择器，能记住上次的目录；Safari 等不支持的浏览器用普通文件框
async function pickFile() {
  if (!window.showOpenFilePicker) {
    fileInput.click();
    return;
  }
  let handle;
  try {
    [handle] = await window.showOpenFilePicker({ id: PICKER_ID, startIn: 'downloads', types: SPREADSHEET_TYPES });
  } catch (error) {
    if (error.name === 'AbortError') return; // 用户取消了选择
    showStatus(`打开文件选择器失败：${error.message}`, 'error');
    console.error(error);
    return;
  }
  await handleFile(await handle.getFile());
}

async function fetchFromLink() {
  let url;
  try {
    url = parseExportLink(linkInput.value);
  } catch (error) {
    showStatus(error.message, 'error');
    return;
  }
  linkInput.value = url;
  const fetchButton = $('fetch-button');
  fetchButton.disabled = true;
  showStatus('正在下载…', 'loading');
  try {
    let response;
    try {
      response = await fetch(url);
    } catch (error) {
      // 链接失效时接口会跳转到不允许跨域的网页，浏览器只报网络错误
      throw new Error(`下载不了，链接可能已失效，或网络不通（${error.message}）`);
    }
    if (!response.ok) throw new Error(`接口返回 ${response.status}，链接可能已失效`);
    const bytes = new Uint8Array(await response.arrayBuffer());
    if (!isSpreadsheet(bytes)) {
      throw new Error(`下载到的不是 Excel 文件（${response.headers.get('content-type') || '未知类型'}），链接可能已失效`);
    }
    loadWorkbook(bytes, LINK_SOURCE_NAME);
  } catch (error) {
    showStatus(`从链接读取失败：${error.message}`, 'error');
    console.error(error);
  } finally {
    fetchButton.disabled = false;
  }
}

function buildOutputFile() {
  const day = selectedDay();
  const bytes = XLSX.write(buildOutputWorkbook(currentRecords, day), { type: 'array', bookType: 'xlsx' });
  return new File([bytes], outputFilename(day), { type: XLSX_MIME });
}

function downloadOutput() {
  const outputFile = outputFileOrReport();
  if (!outputFile) return;
  const link = document.createElement('a');
  link.href = URL.createObjectURL(outputFile);
  link.download = outputFile.name;
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), BLOB_URL_LIFETIME_MS);
}

// 不能是 async，也不能在 share 之前 await，否则 Safari 认为不是用户点击触发的
function shareOutput() {
  const outputFile = outputFileOrReport();
  if (!outputFile) return;
  const shareData = { files: [outputFile] };
  if (!navigator.canShare(shareData)) {
    showStatus('当前浏览器不支持分享文件，请用下载', 'error');
    return;
  }
  navigator.share(shareData).catch((error) => {
    if (error.name === 'AbortError') return; // 用户关闭了分享面板
    showStatus(`分享失败（${error.name}）：${error.message}`, 'error');
    console.error(error);
  });
}

function setupDragAndDrop() {
  const dropZone = $('drop-zone');
  ['dragenter', 'dragover'].forEach((type) => dropZone.addEventListener(type, (event) => {
    event.preventDefault();
    dropZone.classList.add('dragging');
  }));
  ['dragleave', 'drop'].forEach((type) => dropZone.addEventListener(type, () => dropZone.classList.remove('dragging')));
  dropZone.addEventListener('drop', (event) => {
    event.preventDefault();
    const file = event.dataTransfer.files[0];
    if (file) handleFile(file);
  });
}

$('pick-button').addEventListener('click', pickFile);
fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) handleFile(fileInput.files[0]);
  fileInput.value = ''; // 允许重复选择同一个文件
});
dateInput.addEventListener('change', prepareOutput);
linkForm.addEventListener('submit', (event) => {
  event.preventDefault();
  fetchFromLink();
});
$('download-button').addEventListener('click', downloadOutput);
$('share-button').addEventListener('click', shareOutput);

setupDragAndDrop();
if (window.showOpenFilePicker) $('pick-hint').textContent = '会记住上次选文件的目录，也可以把文件拖到这里';
if (navigator.clipboard && navigator.clipboard.readText) {
  const pasteButton = $('paste-button');
  pasteButton.hidden = false;
  pasteButton.addEventListener('click', async () => {
    try {
      linkInput.value = await navigator.clipboard.readText();
    } catch (error) {
      showStatus(`读取剪贴板失败：${error.message}，请长按输入框手动粘贴`, 'error');
      return;
    }
    fetchFromLink();
  });
}
// iOS Safari 支持分享文件，可以直接发到微信或存到"文件"
if (navigator.canShare) $('share-button').hidden = false;
