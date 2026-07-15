import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

const backend = fs.readFileSync(path.resolve(import.meta.dirname, '../apps-script/Code.gs'), 'utf8');
const utilities = {
  base64Decode(value) {
    return [...Buffer.from(value, 'base64')].map(byte => byte > 127 ? byte - 256 : byte);
  }
};
const { validateFaviconData_ } = new Function('Utilities', `${backend}\nreturn { validateFaviconData_ };`)(utilities);

function pngHeader(width = 64, height = 64, totalBytes = 33) {
  const bytes = Buffer.alloc(totalBytes);
  Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]).copy(bytes, 0);
  bytes.writeUInt32BE(13, 8);
  bytes.write('IHDR', 12, 'ascii');
  bytes.writeUInt32BE(width, 16);
  bytes.writeUInt32BE(height, 20);
  bytes[24] = 8;
  bytes[25] = 6;
  return `data:image/png;base64,${bytes.toString('base64')}`;
}

test('서버는 64×64 PNG 파비콘과 기본값 복구만 허용한다', () => {
  const valid = pngHeader();
  assert.equal(validateFaviconData_(valid), valid);
  assert.equal(validateFaviconData_(''), '');
  assert.throws(() => validateFaviconData_(valid.replace('image/png', 'image/svg+xml')), /PNG/);
  assert.throws(() => validateFaviconData_(pngHeader(32, 64)), /64픽셀/);
  assert.throws(() => validateFaviconData_(pngHeader(64, 64, 32769)), /32KB/);
});
