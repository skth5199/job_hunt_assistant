#!/usr/bin/env node
/**
 * sync-tracker.mjs — Single source of truth: data/Application_Tracker.xlsx
 *
 * On each run:
 * 1. Reads ESG&Data + Project Management sheets from the xlsx
 * 2. Reads evaluation data from existing applications.md + pending TSVs + cv-additions
 * 3. Any evaluation/cv entry NOT already in xlsx is appended to ESG&Data
 * 4. Overlays Score, PDF, Report, Notes, CV Version onto all matching rows in-place
 * 5. Rebuilds data/applications.md from the full xlsx (no md-only concept)
 *
 * Run: node sync-tracker.mjs [--dry-run]
 */

import { readFileSync, writeFileSync, existsSync, readdirSync, mkdirSync, renameSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';
import XLSX from 'xlsx';

const ROOT = dirname(fileURLToPath(import.meta.url));
const XLSX_FILE    = join(ROOT, 'data/Application_Tracker.xlsx');
const APPS_FILE    = join(ROOT, 'data/applications.md');
const ADDITIONS_DIR    = join(ROOT, 'batch/tracker-additions');
const CV_ADDITIONS_DIR = join(ROOT, 'batch/cv-additions');
const MERGED_DIR   = join(ADDITIONS_DIR, 'merged');
const DRY_RUN = process.argv.includes('--dry-run');

const EVAL_COLS = ['Score', 'PDF', 'Report', 'Notes', 'CV Version'];

// ─── Helpers ──────────────────────────────────────────────────────────────────

function norm(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function companyFuzzy(a, b) {
  const na = norm(a), nb = norm(b);
  return na === nb || na.includes(nb) || nb.includes(na);
}

function roleFuzzy(a, b) {
  if (norm(a) === norm(b)) return true;
  const wa = a.toLowerCase().split(/\s+/).filter(w => w.length > 3);
  const wb = b.toLowerCase().split(/\s+/).filter(w => w.length > 3);
  return wa.filter(w => wb.some(x => x.includes(w) || w.includes(x))).length >= 2;
}

function excelDateToISO(val) {
  if (!val) return '';
  if (typeof val === 'string' && /^\d{4}-\d{2}-\d{2}/.test(val)) return val.slice(0, 10);
  if (typeof val === 'number') {
    const d = XLSX.SSF.parse_date_code(val);
    if (d) return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
  }
  if (val instanceof Date) return val.toISOString().slice(0, 10);
  return String(val).trim();
}

function normalizeRowKeys(row) {
  const out = {};
  for (const [k, v] of Object.entries(row)) out[k.trim()] = v;
  return out;
}

function c(v) { return String(v || '').trim().replace(/\|/g, '\\|'); }

function extractFilename(s) {
  if (!s) return '';
  const m = String(s).match(/\((?:.*\/)?([^)]+)\)/);
  return m ? m[1] : String(s).replace(/^.*\//, '').trim();
}

// ─── Read xlsx ────────────────────────────────────────────────────────────────

if (!existsSync(XLSX_FILE)) {
  console.error(`❌ Not found: ${XLSX_FILE}`);
  process.exit(1);
}

const wb = XLSX.readFile(XLSX_FILE);

function readSheetRows(name) {
  const ws = wb.Sheets[name];
  if (!ws) { console.warn(`⚠️  Sheet not found: ${name}`); return []; }
  return XLSX.utils.sheet_to_json(ws, { defval: '' }).map(normalizeRowKeys);
}

const esgRaw = readSheetRows('ESG&Data');
const pmRaw  = readSheetRows('Project Management');

function fromXlsx(row, sheet) {
  return {
    company:     String(row['Company Name'] || '').trim(),
    role:        String(row['Role Title'] || '').trim(),
    date:        excelDateToISO(row['Date applied']),
    location:    String(row['Location'] || '').trim(),
    source:      String(row['Source'] || '').trim(),
    roleType:    String(row['Role type'] || '').trim(),
    portal:      String(row['Application portal'] || '').trim(),
    easyApply:   String(row['Easy Apply'] || '').trim(),
    workType:    String(row['Work Type'] || '').trim(),
    status:      String(row['Rejected'] || row['Status'] || '').trim(),
    referral:    String(row['Referal Status'] || '').trim(),
    coverLetter: String(row['Cover Letter'] || '').trim(),
    jd:          String(row['Job Description'] || '').trim(),
    cvVersion:   String(row['CV Version'] || '').trim(),
    score:       String(row['Score'] || '').trim(),
    pdf:         String(row['PDF'] || '').trim(),
    report:      String(row['Report'] || '').trim(),
    notes:       String(row['Notes'] || '').trim(),
    _sheet: sheet,
  };
}

const xlsxEntries = [
  ...esgRaw.map(r => fromXlsx(r, 'ESG&Data')),
  ...pmRaw.map(r  => fromXlsx(r, 'PM')),
].filter(e => e.company);

console.log(`📊 xlsx: ESG&Data=${esgRaw.filter(r => r['Company Name']).length}, PM=${pmRaw.filter(r => r['Company Name']).length} → ${xlsxEntries.length} rows`);

// ─── Collect all evaluation/cv data from every source ─────────────────────────
// Key: norm(company)+'::'+norm(role) → { company, role, date, status, score, pdf, report, notes, cvVersion }

const sourceMap = new Map();

function mergeIntoSource(company, role, data) {
  const key = norm(company) + '::' + norm(role);
  if (!sourceMap.has(key)) sourceMap.set(key, { company, role });
  const existing = sourceMap.get(key);
  for (const [k, v] of Object.entries(data)) {
    if (v && !existing[k]) existing[k] = v;
  }
}

// 1. Existing applications.md (eval columns only)
if (existsSync(APPS_FILE)) {
  const lines = readFileSync(APPS_FILE, 'utf-8').split('\n');
  for (const line of lines) {
    if (!line.startsWith('|') || line.includes('---')) continue;
    const parts = line.split('|').map(s => s.trim());
    if (isNaN(parseInt(parts[1])) || parseInt(parts[1]) === 0) continue;
    let company, role, date, status, score, pdf, report, notes;
    if (parts.length >= 21) {
      [, , date, company, role] = parts;
      status = parts[11]; score = parts[16]; pdf = parts[17];
      report = extractFilename(parts[18]); notes = parts[19];
    } else if (parts.length >= 11) {
      [, , date, company, role, score, status, pdf] = parts;
      report = extractFilename(parts[8]); notes = parts[9];
    } else continue;
    if ((score || report) && company && role) {
      mergeIntoSource(company, role, { date, status, score, pdf, report, notes });
    }
  }
}

// 2. Pending TSV additions
const tsvFiles = [];
if (existsSync(ADDITIONS_DIR)) {
  for (const file of readdirSync(ADDITIONS_DIR).filter(f => f.endsWith('.tsv'))) {
    const raw = readFileSync(join(ADDITIONS_DIR, file), 'utf-8').trim();
    if (!raw) continue;
    const parts = raw.startsWith('|')
      ? raw.split('|').map(s => s.trim()).filter(Boolean)
      : raw.split('\t').map(s => s.trim());
    if (parts.length < 8) { console.warn(`⚠️  Skipping malformed TSV: ${file}`); continue; }
    const company = parts[2], role = parts[3];
    if (!company || !role) continue;
    const col4 = parts[4], col5 = parts[5];
    const c4score  = /^\d+\.?\d*\/5$/.test(col4);
    const c4status = /^(evaluated|applied|responded|interview|offer|rejected|discarded|skip)/i.test(col4);
    let status, score;
    if (c4score && !c4status) { status = col5; score = col4; }
    else if (/^\d+\.?\d*\/5$/.test(col5) && !c4score) { status = col4; score = col5; }
    else { status = col4; score = col5; }
    mergeIntoSource(company, role, {
      date: parts[1], status, score,
      pdf: parts[6], report: extractFilename(parts[7]), notes: parts[8] || '',
    });
    tsvFiles.push(file);
  }
  if (tsvFiles.length > 0) console.log(`📥 ${tsvFiles.length} TSV addition(s) loaded`);
}

// 3. Pending CV additions
const cvAdditionFiles = [];
if (existsSync(CV_ADDITIONS_DIR)) {
  for (const file of readdirSync(CV_ADDITIONS_DIR).filter(f => f.endsWith('.tsv'))) {
    const raw = readFileSync(join(CV_ADDITIONS_DIR, file), 'utf-8').trim();
    if (!raw) continue;
    const parts = raw.split('\t').map(s => s.trim());
    if (parts.length < 3) { console.warn(`⚠️  Skipping malformed cv-addition: ${file}`); continue; }
    const [company, role, cvVersion] = parts;
    if (!company || !role || !cvVersion) continue;
    mergeIntoSource(company, role, { cvVersion });
    cvAdditionFiles.push(file);
  }
  if (cvAdditionFiles.length > 0) console.log(`🖨️  ${cvAdditionFiles.length} CV addition(s) loaded`);
}

// ─── Identify which source entries are missing from xlsx → append ─────────────

const toAppend = [];
for (const [, src] of sourceMap) {
  const inXlsx = xlsxEntries.some(e =>
    companyFuzzy(e.company, src.company) && roleFuzzy(e.role, src.role)
  );
  if (!inXlsx) toAppend.push(src);
}

if (toAppend.length > 0) {
  console.log(`➕ ${toAppend.length} new entr${toAppend.length === 1 ? 'y' : 'ies'} to append to ESG&Data:`);
  toAppend.forEach(e => console.log(`   ${e.company} — ${e.role}`));
}

// ─── Lookup helpers ───────────────────────────────────────────────────────────

function lookupSource(company, role) {
  const exact = norm(company) + '::' + norm(role);
  if (sourceMap.has(exact)) return sourceMap.get(exact);
  for (const [, v] of sourceMap) {
    if (companyFuzzy(company, v.company) && roleFuzzy(role, v.role)) return v;
  }
  return null;
}

// ─── Update xlsx sheets in-place ─────────────────────────────────────────────

function getOrAddEvalCols(ws) {
  const range = XLSX.utils.decode_range(ws['!ref']);
  const colIndex = {};
  for (let C = range.s.c; C <= range.e.c; C++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c: C })];
    if (cell?.v != null) colIndex[String(cell.v).trim()] = C;
  }
  let nextCol = range.e.c + 1;
  for (const name of EVAL_COLS) {
    if (colIndex[name] == null) {
      ws[XLSX.utils.encode_cell({ r: 0, c: nextCol })] = { v: name, t: 's' };
      colIndex[name] = nextCol++;
    }
  }
  if (nextCol - 1 > range.e.c) {
    range.e.c = nextCol - 1;
    ws['!ref'] = XLSX.utils.encode_range(range);
  }
  return colIndex;
}

function writeEvalCells(ws, R, colIndex, src) {
  const write = (col, val) => {
    if (!val) return;
    ws[XLSX.utils.encode_cell({ r: R, c: colIndex[col] })] = { v: val, t: 's' };
  };
  write('Score',      src.score);
  write('PDF',        src.pdf);
  write('Report',     src.report);
  write('Notes',      src.notes);
  write('CV Version', src.cvVersion);
}

function updateSheetInPlace(ws, sheetLabel) {
  if (!ws || !ws['!ref']) return 0;
  const colIndex = getOrAddEvalCols(ws);
  const companyCol = colIndex['Company Name'];
  const roleCol    = colIndex['Role Title'];
  if (companyCol == null || roleCol == null) return 0;
  const range = XLSX.utils.decode_range(ws['!ref']);
  let updated = 0;
  for (let R = 1; R <= range.e.r; R++) {
    const company = String(ws[XLSX.utils.encode_cell({ r: R, c: companyCol })]?.v || '').trim();
    const role    = String(ws[XLSX.utils.encode_cell({ r: R, c: roleCol })]?.v || '').trim();
    if (!company) continue;
    const src = lookupSource(company, role);
    if (!src) continue;
    writeEvalCells(ws, R, colIndex, src);
    updated++;
  }
  return updated;
}

function appendNewEntries(ws, entries) {
  if (!ws || !ws['!ref'] || entries.length === 0) return 0;
  const colIndex = getOrAddEvalCols(ws);
  const range = XLSX.utils.decode_range(ws['!ref']);

  // Max S No. for auto-increment
  const sNoCol = colIndex['S No.'];
  let maxSNo = 0;
  if (sNoCol != null) {
    for (let R = 1; R <= range.e.r; R++) {
      const v = parseInt(ws[XLSX.utils.encode_cell({ r: R, c: sNoCol })]?.v);
      if (!isNaN(v) && v > maxSNo) maxSNo = v;
    }
  }

  const fieldMap = {
    'Company Name': 'company', 'Role Title': 'role',
    'Date applied': 'date',   'Rejected': 'status',
  };

  let appended = 0;
  for (const entry of entries) {
    const R = range.e.r + 1 + appended;
    if (sNoCol != null) {
      ws[XLSX.utils.encode_cell({ r: R, c: sNoCol })] = { v: maxSNo + 1 + appended, t: 'n' };
    }
    for (const [header, field] of Object.entries(fieldMap)) {
      const C = colIndex[header];
      const val = entry[field];
      if (C == null || !val) continue;
      ws[XLSX.utils.encode_cell({ r: R, c: C })] = { v: val, t: 's' };
    }
    writeEvalCells(ws, R, colIndex, entry);
    appended++;
  }

  if (appended > 0) {
    range.e.r += appended;
    ws['!ref'] = XLSX.utils.encode_range(range);
  }
  return appended;
}

if (!DRY_RUN) {
  // Remove legacy Evaluations sheet if present
  if (wb.SheetNames.includes('Evaluations')) {
    wb.SheetNames.splice(wb.SheetNames.indexOf('Evaluations'), 1);
    delete wb.Sheets['Evaluations'];
  }

  const esgUpdated  = updateSheetInPlace(wb.Sheets['ESG&Data'], 'ESG&Data');
  const pmUpdated   = updateSheetInPlace(wb.Sheets['Project Management'], 'PM');
  const esgAppended = appendNewEntries(wb.Sheets['ESG&Data'], toAppend);

  XLSX.writeFile(wb, XLSX_FILE);
  console.log(`✅ xlsx — ESG&Data: ${esgUpdated} updated, ${esgAppended} appended; PM: ${pmUpdated} updated`);
} else {
  console.log('(dry-run) Would update both sheets in-place and append:');
  toAppend.forEach(e => console.log(`  ${e.company} — ${e.role}`));
}

// ─── Re-read xlsx and build applications.md ───────────────────────────────────
// Re-read so appended rows are included in the markdown output

const wbFresh = DRY_RUN ? wb : XLSX.readFile(XLSX_FILE);
const allEntries = [
  ...XLSX.utils.sheet_to_json(wbFresh.Sheets['ESG&Data'],          { defval: '' }).map(normalizeRowKeys).map(r => fromXlsx(r, 'ESG&Data')),
  ...XLSX.utils.sheet_to_json(wbFresh.Sheets['Project Management'], { defval: '' }).map(normalizeRowKeys).map(r => fromXlsx(r, 'PM')),
].filter(e => e.company);

const HEADER = `# Applications Tracker

| # | Date | Company | Role | Location | Source | Role Type | Portal | Easy Apply | Work Type | Status | Referral | Cover Letter | JD | CV Version | Score | PDF | Report | Notes |
|---|------|---------|------|----------|--------|-----------|--------|------------|-----------|--------|----------|--------------|----|-----------:|-------|-----|--------|-------|`;

const mdRows = allEntries.map((e, i) =>
  `| ${i + 1} | ${c(e.date)} | ${c(e.company)} | ${c(e.role)} | ${c(e.location)} | ${c(e.source)} | ${c(e.roleType)} | ${c(e.portal)} | ${c(e.easyApply)} | ${c(e.workType)} | ${c(e.status)} | ${c(e.referral)} | ${c(e.coverLetter)} | ${c(e.jd)} | ${c(e.cvVersion)} | ${c(e.score)} | ${c(e.pdf)} | ${c(e.report)} | ${c(e.notes)} |`
);

if (!DRY_RUN) {
  writeFileSync(APPS_FILE, HEADER + '\n' + mdRows.join('\n') + '\n');
  console.log(`✅ data/applications.md rebuilt — ${allEntries.length} entries`);
} else {
  console.log(`(dry-run) Would write ${allEntries.length} entries to data/applications.md`);
}

// ─── Move processed files to merged/ ─────────────────────────────────────────

if (!DRY_RUN) {
  if (tsvFiles.length > 0) {
    mkdirSync(MERGED_DIR, { recursive: true });
    tsvFiles.forEach(f => renameSync(join(ADDITIONS_DIR, f), join(MERGED_DIR, f)));
    console.log(`📦 Moved ${tsvFiles.length} TSV file(s) to merged/`);
  }
  if (cvAdditionFiles.length > 0) {
    const cvMerged = join(CV_ADDITIONS_DIR, 'merged');
    mkdirSync(cvMerged, { recursive: true });
    cvAdditionFiles.forEach(f => renameSync(join(CV_ADDITIONS_DIR, f), join(cvMerged, f)));
    console.log(`📦 Moved ${cvAdditionFiles.length} cv-addition file(s) to merged/`);
  }
}

console.log(`\n📊 Total: ${allEntries.length} entries`);
if (DRY_RUN) console.log('(dry-run — no files written)');
