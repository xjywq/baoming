import test from 'node:test';
import assert from 'node:assert';
import worker from '../src/worker.js';

async function fetchVia(targetUrl) {
  const request = new Request(`https://baoming.example/api/fetch?url=${encodeURIComponent(targetUrl)}`);
  const response = await worker.fetch(request, {});
  return { status: response.status, body: await response.json() };
}

test('拒绝非报名小程序的导出链接', async () => {
  const rejected = [
    'https://example.com/mini/enroll/v2/export?id=1',
    'http://api-xcx.qunsou.co/mini/enroll/v2/export?id=1',
    'https://api-xcx.qunsou.co.evil.com/mini/enroll/v2/export?id=1',
    'https://api-xcx.qunsou.co/mini/enroll/v2/other?id=1',
    'https://api-xcx.qunsou.co/mini/enroll/v2/export',
    'https://api-xcx.qunsou.co/mini/enroll/v2/export?id=',
  ];
  for (const url of rejected) {
    const { status, body } = await fetchVia(url);
    assert.strictEqual(status, 400, url);
    assert.match(body.error, /只支持报名小程序的导出链接/, url);
  }
});

test('无效链接给出原文', async () => {
  const { status, body } = await fetchVia('不是链接');
  assert.strictEqual(status, 400);
  assert.match(body.error, /不是有效的链接：不是链接/);
});
