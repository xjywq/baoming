// 报名小程序导出链接：浏览器直接下载（接口允许跨域），只放行这一个接口
const EXPORT_ORIGIN = 'https://api-xcx.qunsou.co';
const EXPORT_PATH = '/mini/enroll/v2/export';
const EXPORT_ID_PARAM = 'id';
const URL_IN_TEXT_PATTERN = /https?:\/\/[^\s"'<>，。；]+/;

// 从粘贴的文字里取出导出链接；不是小程序导出链接就报错
function parseExportLink(text) {
  const match = text.match(URL_IN_TEXT_PATTERN);
  if (!match) throw new Error('没有找到链接');
  const url = new URL(match[0]);
  if (url.origin !== EXPORT_ORIGIN || url.pathname !== EXPORT_PATH || !url.searchParams.get(EXPORT_ID_PARAM)) {
    throw new Error(`只支持报名小程序的导出链接（${EXPORT_ORIGIN}${EXPORT_PATH}?${EXPORT_ID_PARAM}=...），收到的是 ${url.origin}${url.pathname}`);
  }
  return url.href;
}

if (typeof module !== 'undefined') {
  module.exports = { parseExportLink };
}
