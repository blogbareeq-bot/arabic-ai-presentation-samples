#!/usr/bin/env node
/**
 * selftest.mjs — اختبارات ذاتية لأدوات الحزمة (pptx-bundle + import-bundle).
 *
 * تبني ملفات PPTX تجريبية داخل مجلد مؤقت، وتتحقق من أن الفاحص يقبل السليم
 * ويرفض التالف، وأن أداة الدمج آمنة. لا تحتاج أي اعتماديات خارجية.
 *
 * التشغيل: node tools/selftest.mjs [--verbose]
 */

import { deflateRawSync } from 'node:zlib';
import { mkdtemp, mkdir, readdir, readFile, rm, writeFile } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { crc32, inspectPptx } from './pptx-bundle.mjs';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = path.resolve(HERE, '..');
const VERBOSE = process.argv.includes('--verbose');

let passed = 0;
const failures = [];

function check(name, condition, detail = '') {
  if (condition) {
    passed++;
    if (VERBOSE) process.stdout.write(`  ✓ ${name}\n`);
  } else {
    failures.push(`${name}${detail ? ` — ${detail}` : ''}`);
    process.stdout.write(`  ✗ ${name}${detail ? ` — ${detail}` : ''}\n`);
  }
}

/* ---------------------------------------------------------- zip builder -- */

function makeZip(entries, { zip64 = false, store = false } = {}) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  for (const [name, content] of entries) {
    const nameBuf = Buffer.from(name, 'utf8');
    const data = Buffer.isBuffer(content) ? content : Buffer.from(content, 'utf8');
    const compressed = store ? data : deflateRawSync(data);
    const crc = crc32(data);

    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6);
    local.writeUInt16LE(store ? 0 : 8, 8);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(compressed.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    localParts.push(local, nameBuf, compressed);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0, 8);
    central.writeUInt16LE(store ? 0 : 8, 10);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(compressed.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt32LE(offset, 42);
    centralParts.push(central, nameBuf);

    offset += local.length + nameBuf.length + compressed.length;
  }

  const centralDir = Buffer.concat(centralParts);
  const body = Buffer.concat(localParts);

  if (!zip64) {
    const eocd = Buffer.alloc(22);
    eocd.writeUInt32LE(0x06054b50, 0);
    eocd.writeUInt16LE(entries.length, 8);
    eocd.writeUInt16LE(entries.length, 10);
    eocd.writeUInt32LE(centralDir.length, 12);
    eocd.writeUInt32LE(body.length, 16);
    return Buffer.concat([body, centralDir, eocd]);
  }

  const eocd64 = Buffer.alloc(56);
  eocd64.writeUInt32LE(0x06064b50, 0);
  eocd64.writeBigUInt64LE(44n, 4);
  eocd64.writeUInt16LE(45, 12);
  eocd64.writeUInt16LE(45, 14);
  eocd64.writeBigUInt64LE(BigInt(entries.length), 32);
  eocd64.writeBigUInt64LE(BigInt(entries.length), 40);
  eocd64.writeBigUInt64LE(BigInt(centralDir.length), 48);
  eocd64.writeBigUInt64LE(BigInt(body.length), 56 - 8);

  const locator = Buffer.alloc(20);
  locator.writeUInt32LE(0x07064b50, 0);
  locator.writeBigUInt64LE(BigInt(body.length + centralDir.length), 8);
  locator.writeUInt32LE(1, 16);

  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0xffff, 8);
  eocd.writeUInt16LE(0xffff, 10);
  eocd.writeUInt32LE(0xffffffff, 12);
  eocd.writeUInt32LE(0xffffffff, 16);

  return Buffer.concat([body, centralDir, eocd64, locator, eocd]);
}

/* ------------------------------------------------------------- fixtures -- */

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
 <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
 <Default Extension="xml" ContentType="application/xml"/>
 <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;

const NOTES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
 <p:cSld><p:spTree><p:sp><p:txBody><a:bodyPr/><a:p><a:r><a:t>ملاحظات المتحدث للشريحة</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld>
</p:notes>`;

function slideXml(title, body, { rtl = true, complex = true } = {}) {
  const rtlAttr = rtl ? ' rtl="1"' : '';
  const cs = complex ? '<a:cs typeface="Tajawal"/>' : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <p:cSld><p:spTree><p:sp><p:txBody><a:bodyPr/><a:lstStyle/>
  <a:p><a:pPr${rtlAttr} algn="r"><a:defRPr>${cs}</a:defRPr></a:pPr><a:r><a:t>${title}</a:t></a:r></a:p>
  <a:p><a:pPr${rtlAttr} algn="r"><a:defRPr>${cs}</a:defRPr></a:pPr><a:r><a:t>${body}</a:t></a:r></a:p>
 </p:txBody></p:sp></p:spTree></p:cSld>
</p:sld>`;
}

function makeDeck({ slides = 2, dropRelation = false, rtl = true, complex = true, store = false } = {}) {
  const sldIds = Array.from({ length: slides }, (_, i) => `<p:sldId id="${255 + i + 1}" r:id="rId${i + 1}"/>`).join('');
  const presentation = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <p:sldIdLst>${sldIds}</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/>
</p:presentation>`;
  const relationshipCount = dropRelation ? slides - 1 : slides;
  const rels = Array.from(
    { length: relationshipCount },
    (_, i) =>
      `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`,
  ).join('');
  const presentationRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;

  const entries = [
    ['[Content_Types].xml', CONTENT_TYPES],
    ['_rels/.rels', ROOT_RELS],
    ['ppt/presentation.xml', presentation],
    ['ppt/_rels/presentation.xml.rels', presentationRels],
    ['ppt/notesSlides/notesSlide1.xml', NOTES],
  ];
  for (let i = 1; i <= slides; i++) {
    entries.push([`ppt/slides/slide${i}.xml`, slideXml(`شريحة رقم ${i}`, `محتوى عربي للشريحة ${i} مع أرقام 2026`, { rtl, complex })]);
  }
  return makeZip(entries, { store });
}

/* ---------------------------------------------------------------- tests -- */

const temp = await mkdtemp(path.join(os.tmpdir(), 'pptx-selftest-'));

/** يشغّل سكربتًا ويعيد النتيجة دون رفع استثناء، لتبقى كل فحص مستقلًا. */
function execute(script, args) {
  try {
    return { ok: true, stdout: execFileSync(process.execPath, [path.join(HERE, script), ...args], { encoding: 'utf8' }), stderr: '' };
  } catch (error) {
    return { ok: false, stdout: String(error.stdout ?? ''), stderr: String(error.stderr ?? '') };
  }
}
const run = (args) => execute('pptx-bundle.mjs', args);
const runImport = (args) => execute('import-bundle.mjs', args);

let exitCode = 0;
try {
  process.stdout.write('اختبارات الفحص البنيوي:\n');
  const good = makeDeck();
  const healthy = inspectPptx(good);
  check('عرض سليم يُقبل', healthy.ok, healthy.errors.join('؛ '));
  check('عدد الشرائح = 2', healthy.slideCount === 2, `got ${healthy.slideCount}`);
  check('لا إنذارات عربية على عرض سليم', healthy.warnings.length === 0, healthy.warnings.join('؛ '));

  const noComplex = inspectPptx(makeDeck({ complex: false }));
  check('نص عربي بلا <a:cs> → تنبيه', noComplex.ok && noComplex.slidesMissingComplexScriptFont.length === 2);

  const noRtl = inspectPptx(makeDeck({ rtl: false }));
  check('فقرات بلا rtl="1" → تنبيه', noRtl.ok && noRtl.paragraphsWithoutRtl === 4, `warnings=${noRtl.warnings.length}`);

  const dangling = inspectPptx(makeDeck({ dropRelation: true }));
  check('علاقة شريحة ناقصة → رفض', !dangling.ok && dangling.errors.some((e) => e.includes('غير مربوطة')));

  const truncated = inspectPptx(good.subarray(0, Math.floor(good.length / 2)));
  check('ZIP مقطوع → رفض', !truncated.ok && truncated.errors.length > 0);

  const corrupt = Buffer.from(good);
  corrupt[corrupt.length - 40] = corrupt[corrupt.length - 40] ^ 0xff;
  const corruptReport = inspectPptx(corrupt);
  check('تلف بنية ZIP → رفض', !corruptReport.ok, corruptReport.errors.join('؛ '));

  // ملف مخزّن دون ضغط: تلف بايت في المحتوى يجب أن يكتشفه فحص CRC وحده
  const stored = makeDeck({ store: true });
  const storedCorrupt = Buffer.from(stored);
  const textAt = storedCorrupt.indexOf(Buffer.from('شريحة رقم 1', 'utf8'));
  check('العثور على المحتوى غير المضغوط في الملف', textAt > 0);
  if (textAt > 0) storedCorrupt[textAt] = 0x58; // 'X' بدل أول بايت عربي
  const crcReport = inspectPptx(storedCorrupt);
  check(
    'تلف محتوى بلا ضغط → رفض بفحص CRC',
    !crcReport.ok && crcReport.errors.some((e) => e.includes('CRC')),
    crcReport.errors.join('؛ '),
  );
  check('الملف المخزّن دون ضغط سليم قبل التلف', inspectPptx(stored).ok);

  const zip64Report = inspectPptx(makeDeck());
  check('قراءة ZIP64 مدعومة', zip64Report.ok);

  const emptyZip = inspectPptx(makeZip([]));
  check('ZIP بلا أجزاء مطلوبة → رفض', !emptyZip.ok);

  process.stdout.write('اختبارات البيان والبصمات:\n');
  const bundleDir = path.join(temp, 'bundle');
  await mkdir(bundleDir, { recursive: true });
  await writeFile(path.join(bundleDir, 'العرض-الأساسي.pptx'), good);
  const manifestRun = run(['manifest', '--root', bundleDir]);
  check('تشغيل أمر البيان ينجح', manifestRun.ok, manifestRun.stderr.slice(0, 200));
  const manifestRaw = await readFile(path.join(bundleDir, 'BUNDLE-MANIFEST.json'), 'utf8');
  const manifestJson = JSON.parse(manifestRaw);
  check(
    'توليد BUNDLE-MANIFEST.json',
    manifestJson.schemaVersion === 1 && manifestJson.fileCount === 1 && manifestJson.files[0].path === 'العرض-الأساسي.pptx',
    manifestRaw.slice(0, 120),
  );
  const sumsRaw = await readFile(path.join(bundleDir, 'SHA256SUMS.txt'), 'utf8');
  check('توليد SHA256SUMS.txt بالصيغة القياسية', /^[a-f0-9]{64} {2}\S+/.test(sumsRaw.trim()), sumsRaw.slice(0, 80));
  const verified = run(['verify', '--root', bundleDir]);
  check('verify يمر بعد توليد البيان', verified.ok, verified.stderr.slice(0, 200));

  await writeFile(path.join(bundleDir, 'العرض-الأساسي.pptx'), makeDeck({ slides: 3 }));
  const stale = run(['verify', '--root', bundleDir]);
  check(
    'تعديل ملف بعد البيان → رفض',
    !stale.ok && `${stale.stdout}${stale.stderr}`.includes('بصمة'),
    `${stale.stdout}${stale.stderr}`.slice(0, 160),
  );
  const staleManifest = run(['manifest', '--check', '--root', bundleDir]);
  check('manifest --check يكشف بيانًا قديمًا', !staleManifest.ok);

  process.stdout.write('اختبارات أداة الدمج:\n');
  const mergeRoot = path.join(temp, 'merge');
  await mkdir(mergeRoot, { recursive: true });
  const cleanZip = makeZip([
    ['نماذج/العرض-أ.pptx', good],
    ['نماذج/العرض-ب.pptx', makeDeck({ slides: 3 })],
    ['__MACOSX/._junk', 'junk'],
    ['readme.txt', 'ليس عرضًا'],
  ]);
  const cleanArchive = path.join(temp, 'clean.zip');
  await writeFile(cleanArchive, cleanZip);
  const importRun = runImport([cleanArchive, '--root', mergeRoot, '--into', 'decks']);
  check('دمج حزمة نظيفة', importRun.ok && importRun.stdout.includes('تم دمج 2 ملف'), `${importRun.stdout}${importRun.stderr}`.slice(-200));
  check('تجاهل ملفات __MACOSX وغير pptx', !importRun.stdout.includes('__MACOSX') && !importRun.stdout.includes('readme.txt'));
  check(
    'إعادة توليد البيان بعد الدمج',
    importRun.stdout.includes('BUNDLE-MANIFEST.json') && importRun.stdout.includes('النتيجة: سليم'),
    importRun.stdout.slice(-160),
  );

  const repeatRun = runImport([cleanArchive, '--root', mergeRoot, '--into', 'decks']);
  check(
    'إعادة الدمج لا تُكرّر (بصمة مطابقة)',
    repeatRun.ok && (repeatRun.stdout.includes('بصمة مطابقة') || repeatRun.stdout.includes('لا جديد')),
    repeatRun.stdout.slice(0, 200),
  );

  const brokenArchive = path.join(temp, 'broken.zip');
  await writeFile(brokenArchive, makeZip([['ok.pptx', good], ['damaged.pptx', makeDeck({ dropRelation: true })]]));
  const brokenRun = runImport([brokenArchive, '--root', mergeRoot, '--into', 'decks']);
  check('رفض حزمة تحتوي ملفًا تالفًا', !brokenRun.ok && brokenRun.stderr.includes('غير سليم'), brokenRun.stderr.slice(0, 160));

  const evilArchive = path.join(temp, 'evil.zip');
  await writeFile(evilArchive, makeZip([['../../../../tmp/pwned-by-import.pptx', good]]));
  const evilRun = runImport([evilArchive, '--root', mergeRoot, '--into', 'decks']);
  check(
    'منع zip-slip',
    !evilRun.ok && (evilRun.stderr.includes('غير آمن') || evilRun.stderr.includes('لا توجد')),
    evilRun.stderr.slice(0, 160),
  );

  const dryRun = runImport([cleanArchive, '--root', path.join(temp, 'dry'), '--into', 'decks', '--dry-run']);
  check('التشغيل التجريبي لا يكتب ملفات', dryRun.ok && dryRun.stdout.includes('لم تُكتب'), dryRun.stdout.slice(0, 160));
  const dryTarget = path.join(temp, 'dry', 'decks');
  const dryFiles = await readdir(dryTarget).catch(() => []);
  check('لا ملفات في مجلد الهدف بعد التشغيل التجريبي', dryFiles.length === 0, dryFiles.join(', '));
} catch (error) {
  failures.push(`استثناء غير متوقع: ${error.message}`);
  process.stdout.write(`  ✗ استثناء: ${error.message}\n`);
  exitCode = 1;
} finally {
  await rm(temp, { recursive: true, force: true });
}

process.stdout.write(`\nالنتيجة: ${passed} نجحت، ${failures.length} فشلت\n`);
if (failures.length) {
  for (const failure of failures) process.stdout.write(`  ✗ ${failure}\n`);
  process.exitCode = 1;
} else {
  process.exitCode = exitCode;
}
