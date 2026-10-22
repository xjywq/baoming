const test = require('node:test');
const assert = require('node:assert');
const { parseExportLink, isSpreadsheet } = require('../public/link.js');

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

test('按文件头识别 Excel 文件', () => {
  assert.ok(isSpreadsheet(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1])));
  assert.ok(isSpreadsheet(new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x14])));
  assert.ok(!isSpreadsheet(new TextEncoder().encode('<!DOCTYPE html>')));
});
