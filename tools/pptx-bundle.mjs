#!/usr/bin/env node
/**
 * pptx-bundle.mjs — التحقق من سلامة حزمة عروض PowerPoint العربية.
 *
 * أداة بلا أي اعتماديات خارجية (Node >= 20) تقرأ ملفات PPTX كمُعرَّفات ZIP
 * وتتحقق من بنية OOXML، وجاهزية العربية (RTL، خط النصوص المركّبة)،
 * وبصمات SHA-256 مقابل MANIFEST.json.
 *
 * الأوامر:
 *   node tools/pptx-bundle.mjs verify   [--root .] [--strict] [--json]
 *   node tools/pptx-bundle.mjs manifest [--root .] [--json]
 *
 * رموز الخروج: 0 = سليم، 1 = أخطاء، 2 = خطأ استخدام.
 */

import { createHash } from 'node:crypto';
import { inflateRawSync } from 'node:zlib';
import { readFile, readdir, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = path.resolve(HERE, '..');

const MANIFEST_NAME = 'BUNDLE-MANIFEST.json';
const SUMS_NAME = 'SHA256SUMS.txt';
const SKIP_DIRS = new Set(['.git', '.github', 'node_modules', 'tools', 'dist', 'build', '.cache', '.venv']);
const REQUIRED_PARTS = ['[Content_Types].xml', '_rels/.rels', 'ppt/presentation.xml'];
const ARABIC_RE = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/;
const SLIDE_RE = /^ppt\/slides\/slide\d+\.xml$/;
const NOTES_RE = /^ppt\/notesSlides\/notesSlide\d+\.xml$/;

/* ------------------------------------------------------------------ ZIP -- */

const CRC_TABLE = (() => {
  const table = new Int32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[i] = c;
  }
  return table;
})();

function crc32(buf) {
  let c = -1;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
}

class ZipError extends Error {}

function findEndOfCentralDirectory(buf) {
  const earliest = Math.max(0, buf.length - 65557);
  for (let i = buf.length - 22; i >= earliest; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) return i;
  }
  return -1;
}

function readZip64Extra(extra, needs) {
  const out = {};
  let p = 0;
  while (p + 4 <= extra.length) {
    const id = extra.readUInt16LE(p);
    const size = extra.readUInt16LE(p + 2);
    const body = extra.subarray(p + 4, p + 4 + size);
    if (id === 0x0001) {
      let q = 0;
      for (const key of needs) {
        if (q + 8 > body.length) break;
        out[key] = Number(body.readBigUInt64LE(q));
        q += 8;
      }
    }
    p += 4 + size;
  }
  return out;
}

/** يقرأ فهرس ZIP المركزي ويعيد قائمة المدخلات. */
export function readZipDirectory(buf) {
  const eocd = findEndOfCentralDirectory(buf);
  if (eocd < 0) throw new ZipError('لم يتم العثور على نهاية فهرس ZIP (ملف تالف أو ليس ZIP).');

  let entryCount = buf.readUInt16LE(eocd + 10);
  let cdOffset = buf.readUInt32LE(eocd + 16);
  let cdSize = buf.readUInt32LE(eocd + 12);

  if (entryCount === 0xffff || cdOffset === 0xffffffff || cdSize === 0xffffffff) {
    const locator = eocd - 20;
    if (locator < 0 || buf.readUInt32LE(locator) !== 0x07064b50) {
      throw new ZipError('فهرس ZIP يشير إلى ZIP64 بدون سجل ZIP64.');
    }
    const z64 = Number(buf.readBigUInt64LE(locator + 8));
    if (z64 < 0 || z64 + 56 > buf.length || buf.readUInt32LE(z64) !== 0x06064b50) {
      throw new ZipError('سجل ZIP64 غير صالح.');
    }
    entryCount = Number(buf.readBigUInt64LE(z64 + 32));
    cdSize = Number(buf.readBigUInt64LE(z64 + 40));
    cdOffset = Number(buf.readBigUInt64LE(z64 + 48));
  }

  if (cdOffset + cdSize > buf.length) throw new ZipError('فهرس ZIP يتجاوز حجم الملف.');

  const entries = [];
  let p = cdOffset;
  for (let i = 0; i < entryCount; i++) {
    if (p + 46 > buf.length || buf.readUInt32LE(p) !== 0x02014b50) {
      throw new ZipError(`مدخل فهرس ZIP رقم ${i + 1} غير صالح.`);
    }
    const flags = buf.readUInt16LE(p + 8);
    const method = buf.readUInt16LE(p + 10);
    const crc = buf.readUInt32LE(p + 16);
    let compressedSize = buf.readUInt32LE(p + 20);
    let uncompressedSize = buf.readUInt32LE(p + 24);
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const commentLen = buf.readUInt16LE(p + 32);
    let localOffset = buf.readUInt32LE(p + 42);
    const name = buf.subarray(p + 46, p + 46 + nameLen).toString('utf8');

    const needs = [];
    if (uncompressedSize === 0xffffffff) needs.push('uncompressedSize');
    if (compressedSize === 0xffffffff) needs.push('compressedSize');
    if (localOffset === 0xffffffff) needs.push('localOffset');
    if (needs.length) {
      const extra = buf.subarray(p + 46 + nameLen, p + 46 + nameLen + extraLen);
      const z64 = readZip64Extra(extra, needs);
      if (z64.uncompressedSize !== undefined) uncompressedSize = z64.uncompressedSize;
      if (z64.compressedSize !== undefined) compressedSize = z64.compressedSize;
      if (z64.localOffset !== undefined) localOffset = z64.localOffset;
    }

    entries.push({ name, method, flags, crc, compressedSize, uncompressedSize, localOffset });
    p += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

/** يستخرج محتوى مدخل ZIP ويتحقق من CRC. */
export function extractEntry(buf, entry) {
  const off = entry.localOffset;
  if (off + 30 > buf.length || buf.readUInt32LE(off) !== 0x04034b50) {
    throw new ZipError(`ترويسة محلية غير صالحة للمدخل «${entry.name}».`);
  }
  const nameLen = buf.readUInt16LE(off + 26);
  const extraLen = buf.readUInt16LE(off + 28);
  const start = off + 30 + nameLen + extraLen;
  const end = start + entry.compressedSize;
  if (end > buf.length) throw new ZipError(`بيانات المدخل «${entry.name}» تتجاوز حجم الملف.`);
  const raw = buf.subarray(start, end);

  let data;
  if (entry.method === 0) data = Buffer.from(raw);
  else if (entry.method === 8) {
    try {
      data = inflateRawSync(raw);
    } catch (error) {
      throw new ZipError(`فك ضغط المدخل «${entry.name}» فشل: ${error.message}`);
    }
  } else throw new ZipError(`طريقة ضغط غير مدعومة (${entry.method}) للمدخل «${entry.name}».`);

  if (data.length !== entry.uncompressedSize) {
    throw new ZipError(`حجم المدخل «${entry.name}» غير مطابق (${data.length} ≠ ${entry.uncompressedSize}).`);
  }
  if (crc32(data) !== entry.crc) throw new ZipError(`CRC غير مطابق للمدخل «${entry.name}» (تلف بيانات).`);
  return data;
}

/* --------------------------------------------------------------- helpers -- */

export const sha256 = (buf) => createHash('sha256').update(buf).digest('hex');

const decodeXml = (value) =>
  value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, code) => String.fromCodePoint(Number(code)))
    .replace(/&amp;/g, '&');

const attrs = (tag) => {
  const out = {};
  for (const [, key, value] of tag.matchAll(/([A-Za-z:_][\w:.-]*)\s*=\s*"([^"]*)"/g)) out[key] = decodeXml(value);
  return out;
};

async function walkPptx(dir, root, found = []) {
  let dirents;
  try {
    dirents = await readdir(dir, { withFileTypes: true });
  } catch {
    return found;
  }
  for (const dirent of dirents.sort((a, b) => a.name.localeCompare(b.name, 'en'))) {
    if (dirent.name.startsWith('.') && dirent.isDirectory() && SKIP_DIRS.has(dirent.name)) continue;
    const full = path.join(dir, dirent.name);
    if (dirent.isDirectory()) {
      if (SKIP_DIRS.has(dirent.name)) continue;
      await walkPptx(full, root, found);
    } else if (/\.pptx$/i.test(dirent.name) && !dirent.name.startsWith('~$') && !dirent.name.startsWith('._')) {
      found.push(path.relative(root, full).split(path.sep).join('/'));
    }
  }
  return found;
}

/* ------------------------------------------------------------ PPTX check -- */

/** يفحص ملف PPTX واحدًا ويعيد تقريرًا مفصّلًا. */
export function inspectPptx(buf) {
  const report = {
    ok: true,
    slideCount: 0,
    notesSlideCount: 0,
    mediaCount: 0,
    declaredSlideCount: 0,
    slidesWithoutArabic: [],
    slidesMissingComplexScriptFont: [],
    paragraphsWithoutRtl: 0,
    errors: [],
    warnings: [],
  };
  const fail = (message) => {
    report.ok = false;
    report.errors.push(message);
  };
  const warn = (message) => report.warnings.push(message);

  let entries;
  try {
    entries = readZipDirectory(buf);
  } catch (error) {
    fail(error.message);
    return report;
  }

  const parts = new Map();
  for (const entry of entries) {
    if (entry.name.endsWith('/')) continue;
    try {
      parts.set(entry.name, extractEntry(buf, entry));
    } catch (error) {
      fail(error.message);
    }
  }

  for (const required of REQUIRED_PARTS) {
    if (!parts.has(required)) fail(`جزء مطلوب مفقود: ${required}`);
  }
  if (report.errors.length) return report;

  const contentTypes = parts.get('[Content_Types].xml').toString('utf8');
  if (!/presentationml\.presentation\.main\+xml/.test(contentTypes)) {
    fail('[Content_Types].xml لا يعرّف نوع العرض التقديمي الرئيسي.');
  }

  const presentation = parts.get('ppt/presentation.xml').toString('utf8');
  if (!/<p:presentation[\s>]/.test(presentation)) fail('ppt/presentation.xml لا يحتوي عنصر <p:presentation>.');

  const slides = [...parts.keys()].filter((name) => SLIDE_RE.test(name)).sort();
  report.slideCount = slides.length;
  report.notesSlideCount = [...parts.keys()].filter((name) => NOTES_RE.test(name)).length;
  report.mediaCount = [...parts.keys()].filter((name) => name.startsWith('ppt/media/')).length;

  const declared = [...presentation.matchAll(/<p:sldId\b[^>]*>/g)].map((m) => attrs(m[0]));
  report.declaredSlideCount = declared.length;
  if (!slides.length) fail('لا توجد شرائح (ppt/slides/slideN.xml) في الملف.');
  if (declared.length !== slides.length) {
    fail(`عدد الشرائح المعلَن (${declared.length}) لا يطابق عدد ملفات الشرائح (${slides.length}).`);
  }

  const relsPath = 'ppt/_rels/presentation.xml.rels';
  if (!parts.has(relsPath)) {
    fail(`جزء العلاقات مفقود: ${relsPath}`);
  } else {
    const rels = parts.get(relsPath).toString('utf8');
    const relById = new Map();
    const slideTargets = new Set();
    for (const [tag] of rels.matchAll(/<Relationship\b[^>]*>/g)) {
      const { Id, Target, Type } = attrs(tag);
      if (Id) relById.set(Id, Target || '');
      if (Type && /\/slide$/.test(Type) && !/slideMaster|slideLayout/.test(Type)) {
        slideTargets.add(path.posix.normalize(path.posix.join('ppt', Target || '')));
      }
    }
    for (const declaredId of declared.map((entry) => entry['r:id']).filter(Boolean)) {
      if (!relById.has(declaredId)) fail(`مرجع علاقة غير معروف في presentation.xml: ${declaredId}`);
    }
    const missing = slides.filter((name) => !slideTargets.has(name));
    const extra = [...slideTargets].filter((name) => !slides.includes(name));
    if (missing.length) fail(`شرائح غير مربوطة بعلاقة: ${missing.join(', ')}`);
    if (extra.length) fail(`علاقات شرائح لا مقابل لها: ${extra.join(', ')}`);
  }

  for (const name of slides) {
    const xml = parts.get(name).toString('utf8');
    if (!/<p:sld[\s>]/.test(xml)) fail(`${name} لا يحتوي العنصر الجذر <p:sld>.`);
    const runs = [...xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => decodeXml(m[1]));
    const arabicRuns = runs.filter((text) => ARABIC_RE.test(text));
    const paragraphs = [...xml.matchAll(/<a:p\b[^>]*>/g)].length;
    const rtlParagraphs = [...xml.matchAll(/<a:pPr\b[^>]*rtl="1"/g)].length;
    report.paragraphsWithoutRtl += Math.max(0, paragraphs - rtlParagraphs);
    if (!xml.includes('<a:cs')) {
      if (arabicRuns.length) report.slidesMissingComplexScriptFont.push(name);
    }
    if (runs.length && !arabicRuns.length) report.slidesWithoutArabic.push(name);
  }

  if (report.slidesMissingComplexScriptFont.length) {
    warn(`${report.slidesMissingComplexScriptFont.length} شريحة بها نص عربي بدون خط نصوص مركّبة <a:cs>.`);
  }
  if (report.slidesWithoutArabic.length) {
    warn(`${report.slidesWithoutArabic.length} شريحة بدون نص عربي (تحقق أنها مقصودة).`);
  }
  if (report.paragraphsWithoutRtl) {
    warn(`${report.paragraphsWithoutRtl} فقرة بدون rtl="1" (قد تكون فقرات لاتينية مقصودة).`);
  }
  if (!report.notesSlideCount) warn('لا توجد صفحات ملاحظات (notesSlides) في الملف.');

  return report;
}

/* --------------------------------------------------------------- command -- */

function parseArgs(argv) {
  const options = { command: argv[0], root: REPO_ROOT, json: false, strict: false, check: false };
  for (let i = 1; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--root') options.root = path.resolve(argv[++i] ?? '.');
    else if (arg.startsWith('--root=')) options.root = path.resolve(arg.slice(7));
    else if (arg === '--json') options.json = true;
    else if (arg === '--strict') options.strict = true;
    else if (arg === '--check') options.check = true;
    else if (arg === '--help' || arg === '-h') options.help = true;
    else return { error: `وسيط غير معروف: ${arg}` };
  }
  return options;
}

async function readJsonIfExists(file) {
  try {
    return JSON.parse(await readFile(file, 'utf8'));
  } catch (error) {
    if (error.code === 'ENOENT') return null;
    throw new Error(`${path.basename(file)} غير صالح: ${error.message}`);
  }
}

function sumsText(files) {
  if (!files.length) return '';
  return `${files.map((file) => `${file.sha256}  ${file.path}`).join('\n')}\n`;
}

async function collectFiles(root) {
  const paths = await walkPptx(root, root);
  const files = [];
  for (const rel of paths) {
    const full = path.join(root, rel);
    const info = await stat(full);
    const buf = await readFile(full);
    const inspection = inspectPptx(buf);
    files.push({
      path: rel,
      bytes: info.size,
      sha256: sha256(buf),
      slides: inspection.slideCount,
      notesSlides: inspection.notesSlideCount,
      media: inspection.mediaCount,
      inspection,
    });
  }
  return files;
}

async function commandManifest(options) {
  const files = await collectFiles(options.root);
  const failed = files.filter((file) => !file.inspection.ok);
  if (failed.length) {
    process.stderr.write(`لا يمكن بناء البيان: ${failed.length} ملف غير سليم.\n`);
    for (const file of failed) {
      for (const error of file.inspection.errors) process.stderr.write(`  - ${file.path}: ${error}\n`);
    }
    return 1;
  }
  const manifest = {
    schemaVersion: 1,
    generator: 'tools/pptx-bundle.mjs',
    fileCount: files.length,
    totalBytes: files.reduce((sum, file) => sum + file.bytes, 0),
    files: files.map(({ path: filePath, bytes, sha256: digest, slides, notesSlides, media }) => ({
      path: filePath,
      bytes,
      sha256: digest,
      slides,
      notesSlides,
      media,
    })),
  };
  const manifestPath = path.join(options.root, MANIFEST_NAME);
  const sumsPath = path.join(options.root, SUMS_NAME);
  if (options.check) {
    const existing = await readJsonIfExists(manifestPath);
    if (!existing) {
      process.stderr.write(`${MANIFEST_NAME} غير موجود — شغّل الأمر بدون --check لتوليده.\n`);
      return 1;
    }
    const same = JSON.stringify(existing) === JSON.stringify(manifest);
    if (!same) process.stderr.write(`${MANIFEST_NAME} لا يطابق الملفات الحالية.\n`);
    return same ? 0 : 1;
  }
  await writeFile(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, 'utf8');
  await writeFile(sumsPath, sumsText(manifest.files), 'utf8');
  if (options.json) {
    process.stdout.write(`${JSON.stringify({ ok: true, manifest: manifestPath, sums: sumsPath, fileCount: files.length, totalBytes: manifest.totalBytes })}\n`);
  } else {
    process.stdout.write(`تم توليد ${MANIFEST_NAME} و${SUMS_NAME} — ${files.length} ملف، ${manifest.totalBytes} بايت.\n`);
  }
  return 0;
}

async function commandVerify(options) {
  const files = await collectFiles(options.root);
  const result = { ok: true, root: options.root, fileCount: files.length, errors: [], warnings: [], files: [] };

  if (!files.length) {
    result.warnings.push('لا توجد ملفات PPTX في النطاق المحدد.');
  }

  for (const file of files) {
    const { inspection } = file;
    if (!inspection.ok) {
      result.ok = false;
      for (const error of inspection.errors) result.errors.push(`${file.path}: ${error}`);
    }
    for (const warning of inspection.warnings) result.warnings.push(`${file.path}: ${warning}`);
    if (options.strict && (inspection.slidesWithoutArabic.length || inspection.slidesMissingComplexScriptFont.length)) {
      result.ok = false;
      result.errors.push(`${file.path}: فشل الفحص الصارم للعربية.`);
    }
    result.files.push({
      path: file.path,
      bytes: file.bytes,
      sha256: file.sha256,
      slides: inspection.slideCount,
      notesSlides: inspection.notesSlideCount,
      media: inspection.mediaCount,
      errors: inspection.errors.length,
      warnings: inspection.warnings.length,
    });
  }

  const manifest = await readJsonIfExists(path.join(options.root, MANIFEST_NAME));
  if (manifest) {
    const expected = new Map((manifest.files ?? []).map((file) => [file.path, file]));
    for (const file of files) {
      const entry = expected.get(file.path);
      if (!entry) {
        result.ok = false;
        result.errors.push(`${file.path}: موجود في المستودع وغير مدرج في ${MANIFEST_NAME}.`);
        continue;
      }
      expected.delete(file.path);
      if (entry.sha256 !== file.sha256) {
        result.ok = false;
        result.errors.push(`${file.path}: بصمة SHA-256 لا تطابق ${MANIFEST_NAME}.`);
      }
      if (entry.bytes !== file.bytes) {
        result.ok = false;
        result.errors.push(`${file.path}: الحجم لا يطابق ${MANIFEST_NAME}.`);
      }
    }
    for (const leftover of expected.keys()) {
      result.ok = false;
      result.errors.push(`${leftover}: مُدرج في ${MANIFEST_NAME} وغير موجود في المستودع.`);
    }
    const sums = await readFile(path.join(options.root, SUMS_NAME), 'utf8').catch(() => null);
    if (sums === null) {
      result.ok = false;
      result.errors.push(`${SUMS_NAME} مفقود.`);
    } else {
      const listed = new Map(
        sums
          .split('\n')
          .filter(Boolean)
          .map((line) => [line.slice(66).trim(), line.slice(0, 64)]),
      );
      for (const file of files) {
        if (listed.get(file.path) !== file.sha256) {
          result.ok = false;
          result.errors.push(`${file.path}: لا يطابق سطر البصمة في ${SUMS_NAME}.`);
        }
      }
    }
  } else {
    result.warnings.push(`${MANIFEST_NAME} غير موجود — تم التحقق من بنية الملفات فقط.`);
  }

  if (options.json) {
    process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
    return result.ok ? 0 : 1;
  }

  const lines = [];
  lines.push(`الحزمة: ${result.root}`);
  lines.push(`ملفات PPTX: ${result.fileCount}`);
  for (const file of result.files) {
    const badge = file.errors ? '✗' : '✓';
    lines.push(`  ${badge} ${file.path} — شرائح: ${file.slides}، ملاحظات: ${file.notesSlides}، وسائط: ${file.media}، ${file.bytes} بايت`);
  }
  if (result.warnings.length) {
    lines.push(`تنبيهات (${result.warnings.length}):`);
    for (const warning of result.warnings) lines.push(`  • ${warning}`);
  }
  if (result.errors.length) {
    lines.push(`أخطاء (${result.errors.length}):`);
    for (const error of result.errors) lines.push(`  ✗ ${error}`);
  }
  lines.push(result.ok ? 'النتيجة: سليم ✅' : 'النتيجة: يوجد أخطاء ❌');
  process.stdout.write(`${lines.join('\n')}\n`);
  return result.ok ? 0 : 1;
}

function usage() {
  return [
    'الاستخدام:',
    '  node tools/pptx-bundle.mjs verify   [--root .] [--strict] [--json]',
    '  node tools/pptx-bundle.mjs manifest [--root .] [--check] [--json]',
  ].join('\n');
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  if (options.error || options.help || !options.command) {
    if (options.error) process.stderr.write(`${options.error}\n`);
    process.stdout.write(`${usage()}\n`);
    return options.error ? 2 : 0;
  }
  if (options.command === 'verify') return commandVerify(options);
  if (options.command === 'manifest') return commandManifest(options);
  process.stderr.write(`أمر غير معروف: ${options.command}\n${usage()}\n`);
  return 2;
}

const invokedDirectly = process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (invokedDirectly) {
  process.exitCode = await main();
}
