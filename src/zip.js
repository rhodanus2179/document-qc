const textDecoder = new TextDecoder('utf-8');
const textEncoder = new TextEncoder();

function u16(view, offset) { return view.getUint16(offset, true); }
function u32(view, offset) { return view.getUint32(offset, true); }

export async function readZip(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  const view = new DataView(arrayBuffer);
  const min = Math.max(0, bytes.length - 65557);
  let eocd = -1;
  for (let i = bytes.length - 22; i >= min; i--) {
    if (u32(view, i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('ZIP終端レコードが見つかりません。OOXMLファイルではない可能性があります。');

  const totalEntries = u16(view, eocd + 10);
  const centralOffset = u32(view, eocd + 16);
  const entries = new Map();
  let ptr = centralOffset;

  for (let index = 0; index < totalEntries; index++) {
    if (u32(view, ptr) !== 0x02014b50) throw new Error('ZIP中央ディレクトリを解析できません。');
    const flags = u16(view, ptr + 8);
    const method = u16(view, ptr + 10);
    const compressedSize = u32(view, ptr + 20);
    const uncompressedSize = u32(view, ptr + 24);
    const nameLen = u16(view, ptr + 28);
    const extraLen = u16(view, ptr + 30);
    const commentLen = u16(view, ptr + 32);
    const localOffset = u32(view, ptr + 42);
    if (compressedSize === 0xffffffff || uncompressedSize === 0xffffffff || localOffset === 0xffffffff) {
      throw new Error('ZIP64形式には現在対応していません。');
    }
    const nameBytes = bytes.subarray(ptr + 46, ptr + 46 + nameLen);
    const name = (flags & 0x800) ? textDecoder.decode(nameBytes) : decodeLegacyName(nameBytes);
    ptr += 46 + nameLen + extraLen + commentLen;
    if (name.endsWith('/')) continue;

    if (u32(view, localOffset) !== 0x04034b50) throw new Error(`ZIPローカルヘッダーが不正です: ${name}`);
    const localNameLen = u16(view, localOffset + 26);
    const localExtraLen = u16(view, localOffset + 28);
    const dataStart = localOffset + 30 + localNameLen + localExtraLen;
    const compressed = bytes.subarray(dataStart, dataStart + compressedSize);
    let data;
    if (method === 0) {
      data = compressed.slice();
    } else if (method === 8) {
      data = await inflateRaw(compressed);
    } else {
      throw new Error(`未対応のZIP圧縮方式です (${method}): ${name}`);
    }
    if (uncompressedSize && data.length !== uncompressedSize) {
      throw new Error(`ZIP展開サイズが一致しません: ${name}`);
    }
    entries.set(normalizePath(name), data);
  }
  return entries;
}

async function inflateRaw(bytes) {
  if (typeof DecompressionStream === 'undefined') {
    throw new Error('このブラウザはDecompressionStreamに対応していません。最新版のChrome/Edge等を使用してください。');
  }
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

function decodeLegacyName(bytes) {
  return textDecoder.decode(bytes);
}

export function normalizePath(path) {
  const parts = [];
  for (const part of path.replace(/\\/g, '/').split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') parts.pop(); else parts.push(part);
  }
  return parts.join('/');
}

export function decodeEntry(entries, name) {
  const bytes = entries.get(normalizePath(name));
  return bytes ? textDecoder.decode(bytes) : null;
}

const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = (c & 1) ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const b of bytes) crc = crcTable[(crc ^ b) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function push16(out, value) { out.push(value & 255, (value >>> 8) & 255); }
function push32(out, value) { out.push(value & 255, (value >>> 8) & 255, (value >>> 16) & 255, (value >>> 24) & 255); }

export function createZipStore(files) {
  const localChunks = [];
  const centralChunks = [];
  let offset = 0;

  for (const [name, content] of Object.entries(files)) {
    const nameBytes = textEncoder.encode(name);
    const data = typeof content === 'string' ? textEncoder.encode(content) : content;
    const crc = crc32(data);
    const local = [];
    push32(local, 0x04034b50); push16(local, 20); push16(local, 0x0800); push16(local, 0);
    push16(local, 0); push16(local, 0); push32(local, crc); push32(local, data.length); push32(local, data.length);
    push16(local, nameBytes.length); push16(local, 0);
    const localBytes = concatBytes([new Uint8Array(local), nameBytes, data]);
    localChunks.push(localBytes);

    const central = [];
    push32(central, 0x02014b50); push16(central, 20); push16(central, 20); push16(central, 0x0800); push16(central, 0);
    push16(central, 0); push16(central, 0); push32(central, crc); push32(central, data.length); push32(central, data.length);
    push16(central, nameBytes.length); push16(central, 0); push16(central, 0); push16(central, 0); push16(central, 0);
    push32(central, 0); push32(central, offset);
    centralChunks.push(concatBytes([new Uint8Array(central), nameBytes]));
    offset += localBytes.length;
  }

  const centralData = concatBytes(centralChunks);
  const localData = concatBytes(localChunks);
  const eocd = [];
  push32(eocd, 0x06054b50); push16(eocd, 0); push16(eocd, 0);
  push16(eocd, centralChunks.length); push16(eocd, centralChunks.length);
  push32(eocd, centralData.length); push32(eocd, localData.length); push16(eocd, 0);
  return concatBytes([localData, centralData, new Uint8Array(eocd)]);
}

function concatBytes(chunks) {
  const length = chunks.reduce((sum, part) => sum + part.length, 0);
  const out = new Uint8Array(length);
  let offset = 0;
  for (const part of chunks) { out.set(part, offset); offset += part.length; }
  return out;
}
