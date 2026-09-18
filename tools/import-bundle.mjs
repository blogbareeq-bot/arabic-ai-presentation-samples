#!/usr/bin/env node
/**
 * import-bundle.mjs — دمج حزمة عروض (ملف ZIP أو مجلد) داخل المستودع.
 *
 * يستخرج ملفات .pptx بأمان (حماية من zip-slip)، يفحص كل ملف قبل الدمج،
 * يمنع التعارض أو التكرار، ثم يعيد توليد البيان والبصمات.
 *
 * الأوامر:
 *   node tools/import-bundle.mjs <حزمة.zip|مجلد> [--into decks] [--dry-run] [--force]
 *
 * رموز الخروج: 0 = نجاح، 1 = رفض، 2 = خطأ استخدام.
 */

import { execFileSync } from 'node:child_process';
import { mkdir, readFile, readdir, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { extractEntry, inspectPptx, readZipDirectory, sha256 } from './pptx-bundle.mjs';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = path.resolve(HERE, '..');
const DEFAULT_INTO = 'decks';
const IGNORED_DIR = /(^|\/)(__MACOSX|\.DS_Store|Thumbs\.db)(\/|$)/i;

const isDeckName = (name) => /\.pptx$/i.test(name) && !path.basename(name).startsWith('~$') && !path.basename(name).startsWith('._');

/** يتحقق أن المسار نسبي وآمن ولا يهرب من الجذر. */
function safeRelative(name) {
  const cleaned = name.replace(/\\/g, '/');
  if (!cleaned || cleaned.includes('\0')) return null;
  if (cleaned.startsWith('/') || /^[A-Za-z]:/.test(cleaned)) return null;
  const normalized = path.posix.normalize(cleaned);
  if (normalized.startsWith('../') || normalized === '..' || normalized.includes('/../')) return null;
  return normalized;
}

async function sha256File(file) {
  return sha256(await readFile(file));
}

async function listDeckFiles(archivePath) {
  const info = await stat(archivePath);
  if (info.isDirectory()) {
    const found = [];
    const walk = async (dir) => {
      for (const dirent of (await readdir(dir, { withFileTypes: true })).sort((a, b) => a.name.localeCompare(b.name, 'en'))) {
        const full = path.join(dir, dirent.name);
        if (dirent.isDirectory()) await walk(full);
        else if (isDeckName(dirent.name)) {
          found.push({ name: path.relative(archivePath, full).split(path.sep).join('/'), bytes: await readFile(full) });
        }
      }
    };
    await walk(archivePath);
    return found;
  }

  const buf = await readFile(archivePath);
  const entries = readZipDirectory(buf);
  const found = [];
  for (const entry of entries) {
    if (entry.name.endsWith('/')) continue;
    if (IGNORED_DIR.test(entry.name)) continue;
    if (!isDeckName(entry.name)) continue;
    const name = safeRelative(entry.name);
    if (!name) throw new Error(`مسار غير آمن داخل الحزمة: ${entry.name}`);
    found.push({ name, bytes: extractEntry(buf, entry) });
  }
  return found;
}

async function main() {
  const argv = process.argv.slice(2);
  let source = null;
  let into = DEFAULT_INTO;
  let root = REPO_ROOT;
  let dryRun = false;
  let force = false;

  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--into') into = argv[++i] ?? DEFAULT_INTO;
    else if (arg.startsWith('--into=')) into = arg.slice(7);
    else if (arg === '--root') root = path.resolve(argv[++i] ?? '.');
    else if (arg.startsWith('--root=')) root = path.resolve(arg.slice(7));
    else if (arg === '--dry-run') dryRun = true;
    else if (arg === '--force') force = true;
    else if (arg === '--help' || arg === '-h') source = null;
    else if (arg.startsWith('-')) {
      process.stderr.write(`وسيط غير معروف: ${arg}\n`);
      return 2;
    } else if (!source) source = arg;
    else {
      process.stderr.write(`مصدر إضافي غير متوقع: ${arg}\n`);
      return 2;
    }
  }

  if (!source) {
    process.stdout.write('الاستخدام: node tools/import-bundle.mjs <حزمة.zip|مجلد> [--into decks] [--dry-run] [--force]\n');
    return source === null && argv.includes('--help') ? 0 : 2;
  }

  const sourcePath = path.resolve(source);
  let decks;
  try {
    decks = await listDeckFiles(sourcePath);
  } catch (error) {
    process.stderr.write(`تعذّر قراءة الحزمة: ${error.message}\n`);
    return 2;
  }

  if (!decks.length) {
    process.stderr.write('لا توجد ملفات .pptx في المصدر.\n');
    return 1;
  }

  // 1) فحص كل ملف قبل أي كتابة
  const rejected = [];
  for (const deck of decks) {
    const report = inspectPptx(deck.bytes);
    deck.inspection = report;
    deck.sha256 = sha256(deck.bytes);
    if (!report.ok) rejected.push(deck);
  }
  if (rejected.length) {
    process.stderr.write(`رُفضت الحزمة: ${rejected.length} ملف غير سليم.\n`);
    for (const deck of rejected) {
      for (const error of deck.inspection.errors) process.stderr.write(`  - ${deck.name}: ${error}\n`);
    }
    return 1;
  }

  // 2) كشف الملفات الموجودة مسبقًا (نفس البصمة أو تعارض الاسم)
  const targetDir = path.resolve(root, into);
  await mkdir(targetDir, { recursive: true });
  const existingByHash = new Map();
  const existingNames = new Set();
  const walkExisting = async (dir) => {
    for (const dirent of (await readdir(dir, { withFileTypes: true })).sort((a, b) => a.name.localeCompare(b.name, 'en'))) {
      if (dirent.name === '.git' || dirent.name === 'node_modules') continue;
      const full = path.join(dir, dirent.name);
      if (dirent.isDirectory()) await walkExisting(full);
      else if (isDeckName(dirent.name)) {
        existingByHash.set(await sha256File(full), path.relative(root, full).split(path.sep).join('/'));
        existingNames.add(path.basename(dirent.name).toLowerCase());
      }
    }
  };
  await walkExisting(root);

  const plan = [];
  const conflicts = [];
  for (const deck of decks) {
    const base = path.basename(deck.name);
    const duplicate = existingByHash.get(deck.sha256);
    if (duplicate) {
      plan.push({ deck, action: 'skip-duplicate', target: duplicate });
      continue;
    }
    if (existingNames.has(base.toLowerCase()) && !force) {
      conflicts.push({ deck, reason: `يوجد ملف بنفس الاسم بمحتوى مختلف: ${base}` });
      continue;
    }
    plan.push({ deck, action: force || existingNames.has(base.toLowerCase()) ? 'overwrite' : 'add', target: path.posix.join(into, base) });
  }

  if (conflicts.length) {
    process.stderr.write(`تعارض في ${conflicts.length} ملف (استخدم --force للاستبدال):\n`);
    for (const conflict of conflicts) process.stderr.write(`  - ${conflict.deck.name}: ${conflict.reason}\n`);
    return 1;
  }

  // 3) التنفيذ
  const written = [];
  for (const item of plan) {
    if (item.action === 'skip-duplicate') continue;
    if (!dryRun) {
      const destination = path.join(root, item.target);
      await mkdir(path.dirname(destination), { recursive: true });
      await writeFile(destination, item.deck.bytes);
    }
    written.push({ file: item.target, action: item.action, deck: item.deck });
  }

  process.stdout.write(`المصدر: ${sourcePath}\n`);
  process.stdout.write(`ملفات في الحزمة: ${decks.length}\n`);
  for (const item of plan) {
    const label = item.action === 'skip-duplicate' ? 'موجود مسبقًا (بصمة مطابقة)' : item.action === 'overwrite' ? 'استبدال' : 'إضافة';
    process.stdout.write(`  • ${item.deck.name} → ${item.target} [${label}] — شرائح: ${item.deck.inspection.slideCount}\n`);
  }
  if (dryRun) {
    process.stdout.write('تشغيل تجريبي (--dry-run): لم تُكتب أي ملفات.\n');
    return 0;
  }
  if (!written.length) {
    process.stdout.write('لا جديد للإضافة — الحزمة مدمجة مسبقًا.\n');
    return 0;
  }

  // 4) إعادة توليد البيان والبصمات
  const manifestScript = path.join(HERE, 'pptx-bundle.mjs');
  execFileSync(process.execPath, [manifestScript, 'manifest', '--root', root], { stdio: 'inherit' });
  const verifyScript = execFileSync(process.execPath, [manifestScript, 'verify', '--root', root], { encoding: 'utf8' });
  process.stdout.write(`${verifyScript}\n`);
  process.stdout.write(`تم دمج ${written.length} ملف.\n`);
  return 0;
}

const invokedDirectly = process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (invokedDirectly) process.exitCode = await main();
