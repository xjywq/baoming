const PICKER_ID = 'roster-source'; // 浏览器按此 id 记住上次选文件的目录
const XLS_TYPES = [{ description: '报名表', accept: { 'application/vnd.ms-excel': ['.xls'] } }];
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const LINK_SOURCE_NAME = '导出链接'; // 接口不允许网页读取文件名，日期从表格标题识别
const BLOB_URL_LIFETIME_MS = 60000;
const NO_SLOT_LABEL = '未选时段';

const $ = (id) => document.getElementById(id);
const fileInput = $('file-input');
const linkInput = $('link-input');
const dateInput = $('date-input');
const statusBox = $('status');

let currentRecords = null;

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

function loadWorkbook(bytes, sourceName) {
  const workbook = readXls(bytes);
  currentRecords = extractRecords(workbook);
  const day = detectDay(sourceName, workbook, new Date().getFullYear());
  dateInput.value = day ? formatIsoDate(day) : '';
  renderResult(sourceName, Boolean(day));
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
    loadWorkbook(new Uint8Array(await file.arrayBuffer()), file.name);
  } catch (error) {
    showStatus(`读取「${file.name}」失败：${error.message}`, 'error');
    console.error(error);
  }
}

// Chrome / Edge 用系统文件选择器，能记住上次的目录；其他浏览器直接用 label 关联的原生文件框
async function pickFileWithPicker() {
  let handle;
  try {
    [handle] = await window.showOpenFilePicker({ id: PICKER_ID, startIn: 'downloads', types: XLS_TYPES });
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
    if (!isXlsFile(bytes)) {
      throw new Error(`下载到的不是 .xls 文件（${response.headers.get('content-type') || '未知类型'}），链接可能已失效`);
    }
    loadWorkbook(bytes, LINK_SOURCE_NAME);
  } catch (error) {
    showStatus(`从链接读取失败：${error.message}`, 'error');
    console.error(error);
  } finally {
    fetchButton.disabled = false;
  }
}

// 生成很快（毫秒级），点击时现做，iOS 分享仍在用户点击的有效期内
function buildOutputFile() {
  if (!dateInput.value) {
    showStatus('请先选择活动日期', 'error');
    return null;
  }
  const [year, month, day] = dateInput.value.split('-').map(Number);
  const date = { year, month, day };
  const bytes = writeXlsx(buildOutputSheet(currentRecords, date), OUTPUT_FONTS);
  return new File([bytes], outputFilename(date), { type: XLSX_MIME });
}

function downloadOutput() {
  const outputFile = buildOutputFile();
  if (!outputFile) return;
  const link = document.createElement('a');
  link.href = URL.createObjectURL(outputFile);
  link.download = outputFile.name;
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), BLOB_URL_LIFETIME_MS);
}

function shareOutput() {
  const outputFile = buildOutputFile();
  if (!outputFile) return;
  navigator.share({ files: [outputFile] }).catch((error) => {
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

function handleChosenInputFile() {
  if (fileInput.files[0]) handleFile(fileInput.files[0]);
  fileInput.value = ''; // 允许重复选择同一个文件
}

if (window.showOpenFilePicker) {
  $('pick-button').addEventListener('click', (event) => {
    event.preventDefault(); // 不打开 label 关联的原生文件框
    pickFileWithPicker();
  });
  $('pick-hint').textContent = '会记住上次选文件的目录，也可以把文件拖到这里';
}
fileInput.addEventListener('change', handleChosenInputFile);
// 脚本还没加载完时用户可能已经通过原生文件框选好了文件
handleChosenInputFile();
$('link-form').addEventListener('submit', (event) => {
  event.preventDefault();
  fetchFromLink();
});
$('download-button').addEventListener('click', downloadOutput);
$('share-button').addEventListener('click', shareOutput);
setupDragAndDrop();

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
if (navigator.canShare) $('share-button').hidden = false;
