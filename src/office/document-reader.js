import { parseOfficeBuffer } from '../ooxml.js';

const DEFAULT_SLICE_SIZE = 1024 * 1024;

export async function readCurrentWordDocument(onProgress = () => {}) {
  ensureOfficeWord();
  onProgress(5, 'Word文書のスナップショットを取得しています');

  try {
    const buffer = await getCompressedDocumentBuffer((received, total) => {
      const ratio = total ? received / total : 0;
      onProgress(8 + Math.round(ratio * 24), `Word文書を読み込んでいます (${received}/${total})`);
    });
    const filename = inferDocumentName();
    const model = await parseOfficeBuffer(buffer, filename, onProgress);
    return { model, filename, source: 'compressed-ooxml' };
  } catch (error) {
    if (isOfficeOnline()) {
      throw new Error('Word on the web では現在の文書を .docx として取得できないため、document-qc の完全チェックはまだ利用できません。Windows / Mac 版 Word で開いてください。');
    }
    throw error;
  }
}

export function getOfficeEnvironmentLabel() {
  if (typeof Office === 'undefined') return 'Office 未接続';
  switch (Office.context?.platform) {
    case Office.PlatformType?.PC: return 'Word for Windows';
    case Office.PlatformType?.Mac: return 'Word for Mac';
    case Office.PlatformType?.OfficeOnline: return 'Word on the web';
    case Office.PlatformType?.iOS: return 'Word for iPad';
    default: return 'Word';
  }
}

function getCompressedDocumentBuffer(onSliceProgress) {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(
      Office.FileType.Compressed,
      { sliceSize: DEFAULT_SLICE_SIZE },
      result => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(result.error?.message || 'Word文書を取得できませんでした。'));
          return;
        }
        readAllSlices(result.value, onSliceProgress).then(resolve, reject);
      }
    );
  });
}

async function readAllSlices(file, onSliceProgress) {
  const chunks = [];
  let totalLength = 0;
  try {
    for (let index = 0; index < file.sliceCount; index++) {
      const slice = await getSlice(file, index);
      const bytes = toUint8Array(slice.data);
      chunks.push(bytes);
      totalLength += bytes.byteLength;
      onSliceProgress(index + 1, file.sliceCount);
    }
  } finally {
    await closeFile(file);
  }

  const joined = new Uint8Array(totalLength);
  let offset = 0;
  for (const chunk of chunks) {
    joined.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return joined.buffer;
}

function getSlice(file, index) {
  return new Promise((resolve, reject) => {
    file.getSliceAsync(index, result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new Error(result.error?.message || `Word文書のスライス ${index + 1} を取得できませんでした。`));
    });
  });
}

function closeFile(file) {
  return new Promise(resolve => {
    try {
      file.closeAsync(() => resolve());
    } catch {
      resolve();
    }
  });
}

function toUint8Array(data) {
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (ArrayBuffer.isView(data)) return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  if (Array.isArray(data)) return Uint8Array.from(data);
  throw new Error('Wordから取得した文書データの形式を認識できません。');
}

function inferDocumentName() {
  const url = Office.context?.document?.url || '';
  if (url) {
    try {
      const pathname = new URL(url).pathname;
      const name = decodeURIComponent(pathname.split('/').filter(Boolean).pop() || '');
      if (/\.docx$/i.test(name)) return name;
    } catch {
      const tail = url.split(/[\\/]/).pop() || '';
      if (/\.docx$/i.test(tail)) return tail;
    }
  }
  return 'current-document.docx';
}

function isOfficeOnline() {
  return typeof Office !== 'undefined' && Office.context?.platform === Office.PlatformType?.OfficeOnline;
}

function ensureOfficeWord() {
  if (typeof Office === 'undefined' || !Office.context?.document) throw new Error('Office.js の初期化が完了していません。');
  if (Office.context.host && Office.context.host !== Office.HostType.Word) throw new Error('このアドインは Word 文書で使用してください。');
}
