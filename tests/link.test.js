const test = require('node:test');
const assert = require('node:assert');
const { parseExportLink } = require('../public/link.js');

const VALID_LINK = 'https://api-xcx.qunsou.co/mini/enroll/v2/export?id=abc123';

test('从粘贴的文字里取出导出链接', () => {
  assert.strictEqual(parseExportLink(VALID_LINK), VALID_LINK);
  assert.strictEqual(parseExportLink(`报名表导出：${VALID_LINK}，请查收`), VALID_LINK);
});

test('拒绝非报名小程序的导出链接', () => {
  const rejected = [
    'https://example.com/mini/enroll/v2/export?id=1',
    'http://api-xcx.qunsou.co/mini/enroll/v2/export?id=1',
    'https://api-xcx.qunsou.co.evil.com/mini/enroll/v2/export?id=1',
    'https://api-xcx.qunsou.co/mini/enroll/v2/other?id=1',
    'https://api-xcx.qunsou.co/mini/enroll/v2/export',
    'https://api-xcx.qunsou.co/mini/enroll/v2/export?id=',
  ];
  for (const link of rejected) {
    assert.throws(() => parseExportLink(link), /只支持报名小程序的导出链接/, link);
  }
  assert.throws(() => parseExportLink('没有链接'), /没有找到链接/);
});

