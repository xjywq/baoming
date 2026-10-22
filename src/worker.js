// 静态页面由 assets 直接提供；这里只处理 /api/fetch：代下载报名表链接，绕过浏览器跨域限制
const FETCH_PATH = '/api/fetch';
const MAX_FILE_BYTES = 5 * 1024 * 1024;
// 只代下载报名小程序的导出链接：https://api-xcx.qunsou.co/mini/enroll/v2/export?id=...
const ALLOWED_ORIGIN = 'https://api-xcx.qunsou.co';
const ALLOWED_PATH = '/mini/enroll/v2/export';
const EXPORT_ID_PARAM = 'id';
// 只放行 Excel 文件，防止被当成通用代理
const SPREADSHEET_SIGNATURES = [
  [0xd0, 0xcf, 0x11, 0xe0], // .xls（OLE）
  [0x50, 0x4b, 0x03, 0x04], // .xlsx（zip）
];
const UPSTREAM_USER_AGENT = 'Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/15E148';

function jsonError(message, status) {
  return new Response(JSON.stringify({ error: message }), {
    status,
    headers: { 'content-type': 'application/json; charset=utf-8' },
  });
}

function hasSpreadsheetSignature(bytes) {
  return SPREADSHEET_SIGNATURES.some((signature) => signature.every((byte, i) => bytes[i] === byte));
}

// header 值按 latin1 给出，原始 UTF-8 文件名需要还原
function decodeLatin1AsUtf8(text) {
  const bytes = Uint8Array.from(text, (ch) => ch.charCodeAt(0));
  return new TextDecoder().decode(bytes);
}

function filenameFrom(response, url) {
  const disposition = response.headers.get('content-disposition') || '';
  const encodedMatch = disposition.match(/filename\*\s*=\s*(?:UTF-8|utf-8)''([^;]+)/);
  if (encodedMatch) return decodeURIComponent(encodedMatch[1]);
  const plainMatch = disposition.match(/filename\s*=\s*"?([^";]+)"?/);
  if (plainMatch) return decodeLatin1AsUtf8(plainMatch[1]);
  const lastSegment = url.pathname.split('/').pop();
  return decodeURIComponent(lastSegment || '');
}

async function proxySpreadsheet(rawUrl) {
  let target;
  try {
    target = new URL(rawUrl || '');
  } catch {
    return jsonError(`不是有效的链接：${rawUrl}`, 400);
  }
  if (target.origin !== ALLOWED_ORIGIN || target.pathname !== ALLOWED_PATH || !target.searchParams.get(EXPORT_ID_PARAM)) {
    return jsonError(`只支持报名小程序的导出链接（${ALLOWED_ORIGIN}${ALLOWED_PATH}?${EXPORT_ID_PARAM}=...），收到的是 ${target.origin}${target.pathname}`, 400);
  }

  let upstream;
  try {
    upstream = await fetch(target, { redirect: 'follow', headers: { 'user-agent': UPSTREAM_USER_AGENT } });
  } catch (error) {
    return jsonError(`无法连接 ${target.host}：${error.message}`, 502);
  }
  if (!upstream.ok) {
    return jsonError(`${target.host} 返回 ${upstream.status} ${upstream.statusText}，链接可能已过期或需要登录`, 502);
  }

  const declaredLength = Number(upstream.headers.get('content-length'));
  if (declaredLength > MAX_FILE_BYTES) {
    return jsonError(`文件 ${declaredLength} 字节，超过 ${MAX_FILE_BYTES} 字节上限`, 413);
  }
  const bytes = new Uint8Array(await upstream.arrayBuffer());
  if (bytes.length > MAX_FILE_BYTES) {
    return jsonError(`文件 ${bytes.length} 字节，超过 ${MAX_FILE_BYTES} 字节上限`, 413);
  }
  if (!hasSpreadsheetSignature(bytes)) {
    const contentType = upstream.headers.get('content-type') || '未知类型';
    return jsonError(`链接下载到的不是 Excel 文件（${contentType}），可能是网页而不是文件下载链接`, 415);
  }

  return new Response(bytes, {
    headers: {
      'content-type': 'application/octet-stream',
      'x-filename': encodeURIComponent(filenameFrom(upstream, new URL(upstream.url || target))),
      'cache-control': 'no-store',
    },
  });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (url.pathname === FETCH_PATH) return proxySpreadsheet(url.searchParams.get('url'));
    return env.ASSETS.fetch(request);
  },
};
