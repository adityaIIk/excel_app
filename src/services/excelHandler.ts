import path from 'path';
// eslint-disable-next-line @typescript-eslint/no-require-imports
const XlsxPopulate = require('xlsx-populate') as {
 fromFileAsync: (path: string) => Promise<XlsxWorkbook>;
};

interface XlsxCell {
 value(): unknown;
 value(v: unknown): XlsxCell;
 formula(): string | undefined;
}
interface XlsxRow { cell(col: number): XlsxCell; }
interface XlsxSheet {
 row(row: number): XlsxRow;
 name(): string;
}
interface XlsxWorkbook {
 sheet(name: string): XlsxSheet;
 outputAsync(): Promise<Buffer>;
}
import { ClaimData, PartEntry } from '../types/index';

const TEMPLATE_PATH = path.join(__dirname, '..', '..', 'Format.xlsx');

// Numeric fields — write as actual numbers so Excel doesn't show "stored as text" warning
const NUMERIC_FIELDS: Set<keyof ClaimData> = new Set([
 'policyNumber', 'insuredPhone', 'cubicCapacity', 'seatingCapacity',
]);

// Date fields — write as DD/MM/YYYY text
const DATE_FIELDS: Set<keyof ClaimData> = new Set([
 'insurancePeriodFrom', 'insurancePeriodTo', 'regDate', 'vehicleSaleDate',
 'accidentDate', 'dateOfIntimation', 'dlValidUpto', 'dlIssuedOn',
]);

/**
 * Parse a date string (any common format) into { day, month, year }.
 * Returns null if unparseable.
 */
function parseDateParts(val: string): { day: number; month: number; year: number } | null {
 // DD/MM/YYYY or DD-MM-YYYY
 let m = val.match(/^(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})$/);
 if (m) return { day: +m[1], month: +m[2], year: +m[3] };
 // YYYY-MM-DD or YYYY/MM/DD
 m = val.match(/^(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})$/);
 if (m) return { day: +m[3], month: +m[2], year: +m[1] };
 // DD-Mon-YYYY (e.g. "09-Dec-2025") — strip time portions
 m = val.match(/^(\d{1,2})[\s\-](Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[\s\-](\d{4})/i);
 if (m) {
 const months: Record<string, number> = { jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12 };
 return { day: +m[1], month: months[m[2].toLowerCase().slice(0,3)], year: +m[3] };
 }
 return null;
}

/**
 * Format any date string into DD/MM/YYYY text.
 * We write it as a plain string — like a typist would — so we never
 * create Date objects or modify cell styles. The template formatting stays intact.
 */
function toDateText(dateStr: string): string {
 const parts = parseDateParts(dateStr.trim());
 if (parts) {
 const dd = String(parts.day).padStart(2, '0');
 const mm = String(parts.month).padStart(2, '0');
 return `${dd}/${mm}/${parts.year}`;
 }
 return dateStr; // can't parse — pass through as-is
}

/**
 * Convert a date string to an Excel serial number (integer).
 * This is equivalent to typing the date and pressing Enter —
 * Excel stores the serial and the cell's existing format displays DD/MM/YYYY.
 */
function toExcelDateSerial(dateStr: string): number | string {
 const parts = parseDateParts(dateStr.trim());
 if (!parts) return dateStr; // can't parse — return as text
 // Excel epoch: Dec 30, 1899 = serial 0 (accounts for Lotus 1-2-3 leap year bug)
 const epoch = Date.UTC(1899, 11, 30);
 const target = Date.UTC(parts.year, parts.month - 1, parts.day);
 return Math.round((target - epoch) / 86400000);
}

function norm(text: string): string {
 return text.toLowerCase().replace(/[^a-z0-9]/g, '');
}

function safeCellText(sheet: XlsxSheet, row: number, col: number): string {
 try {
 const v = sheet.row(row).cell(col).value();
 if (v === null || v === undefined) return '';
 return String(v).trim();
 } catch { return ''; }
}

/** Returns true if a cell contains a formula — we must never overwrite these */
function hasFormula(sheet: XlsxSheet, row: number, col: number): boolean {
 try {
 return !!sheet.row(row).cell(col).formula();
 } catch { return false; }
}

/**
 * Scan every cell in a sheet and build a map: normalisedLabel → { row, col } (1-based)
 * Only scans the data entry area to avoid touching calculation sections.
 */
function buildLabelIndex(
 sheet: XlsxSheet,
 maxRow = 70,
 maxCol = 10
): Map<string, { row: number; col: number }> {
 const map = new Map<string, { row: number; col: number }>();
 for (let r = 1; r <= maxRow; r++) {
 for (let c = 1; c <= maxCol; c++) {
 try {
 const text = safeCellText(sheet, r, c);
 if (text.length < 2) continue;
 const n = norm(text);
 if (!map.has(n)) map.set(n, { row: r, col: c });
 } catch { /* skip */ }
 }
 }
 return map;
}

// Fields we must NEVER touch — they contain formulas, pre-filled values, or sheet logic
const DO_NOT_TOUCH_LABELS = new Set(
 [
 'POLICY STATUS', 'COMPULSORY EXCESS', 'AGE OF VEHICLE', 'AGE OF VEH',
 'IMPOSED EXCESS', 'VOLUNTARY EXCESS', 'NO. OF CLAIM(S)', 'NO OF CLAIMS',
 'IMT ENDORSEMENT', '64 VB Compliance',
 ].map(norm)
);

// Label variants for every ClaimData field
const FIELD_LABEL_MAP: Record<keyof ClaimData, string[]> = {
 claimNumber: ['CLAIM NUMBER'],
 policyNumber: ['POLICY NUMBER'],
 insuredName: ['NAME OF INSURED'],
 insuredPhone: ["INSURED'S PHONE NO.", 'INSURED PHONE NO'],
 insuredEmail: [], // never touch — template has its own logic
 vehicleType: ['TYPE OF PRIVATE VEHICLE'],
 financerName: ['FINANCER NAME'],
 insurancePeriodFrom: ['PERIOD OF INSURANCE'],
 insurancePeriodTo: [], // handled alongside insurancePeriodFrom
 idv: ['IDV OF VEHICLE'],
 regNumber: ['Registration Number'],
 regDate: ['Registration Date'],
 vehicleSaleDate: ['Vehicle Sale Date'],
 chassisNumber: ['Chassis Number'],
 engineNumber: ['Engine Number'],
 makeModel: ['Make/Model'],
 colour: ['Colour of Vehicle'],
 cubicCapacity: ['Cubic Capacity'],
 seatingCapacity: ['Seating Capacity/Passenger', 'Seating Capacity'],
 accidentDate: ['Date of accident'],
 accidentPlace: ['Place of accident'],
 dateOfIntimation: ['Date of intimation'],
 estimatedLoss: ['Estimated Loss'],
 causeOfLoss: ['Cause of loss'],
 driverName: ['Driver name'],
 dlNumber: ['Driving licence number'],
 dlValidUpto: ['Valid upto'],
 dlIssuedOn: ['Issued on'],
 dlIssuingAuthority: ['Issuing Authority'],
 licenseType: ['Type of License'],
};

export async function fillAndDownload(claimData: ClaimData, parts: PartEntry[]): Promise<Buffer> {
 // Load the template — xlsx-populate reads the raw xlsx ZIP and only modifies cells we touch
 const workbook = await XlsxPopulate.fromFileAsync(TEMPLATE_PATH);

 // ── DATA ENTRY SHEET ──────────────────────────────────────────────────────
 const wsDE = workbook.sheet('Data Entry');
 if (!wsDE) {
 console.warn('Data Entry sheet not found');
 } else {
 const labelIndex = buildLabelIndex(wsDE);
 console.log(`Data Entry: ${labelIndex.size} labels found`);

 // Build protected cells set from DO_NOT_TOUCH labels only (no broad scanning)
 const protectedCells = new Set<string>();
 for (const [label, pos] of labelIndex.entries()) {
 if (DO_NOT_TOUCH_LABELS.has(label)) {
 for (let offset = 1; offset <= 4; offset++) {
 protectedCells.add(`${pos.row},${pos.col + offset}`);
 }
 }
 }

 for (const [fieldKey, variants] of Object.entries(FIELD_LABEL_MAP)) {
 const key = fieldKey as keyof ClaimData;
 const value = claimData[key];
 if (!value) continue;
 if (variants.length === 0) continue; // field explicitly disabled

 for (const labelText of variants) {
 const lc = labelIndex.get(norm(labelText));
 if (!lc) continue;

 const targetCol = lc.col + 1;

 // Never write to a protected cell (DO_NOT_TOUCH label)
 if (protectedCells.has(`${lc.row},${targetCol}`)) {
 console.log(` ${key} → row ${lc.row} col ${targetCol} SKIPPED (protected)`);
 break;
 }

 // Never overwrite a formula cell — check only the ONE cell we're about to write
 if (hasFormula(wsDE, lc.row, targetCol)) {
 console.log(` ${key} → row ${lc.row} col ${targetCol} SKIPPED (formula)`);
 break;
 }

 if (key === 'insurancePeriodFrom') {
 wsDE.row(lc.row).cell(targetCol).value(toExcelDateSerial(String(value)));
 // Find "To" separator in the same row and write insurancePeriodTo after it
 if (claimData.insurancePeriodTo) {
 for (let offset = 2; offset <= 8; offset++) {
 try {
 const txt = safeCellText(wsDE, lc.row, lc.col + offset);
 if (txt.toLowerCase() === 'to') {
 const toCol = lc.col + offset + 1;
 if (!protectedCells.has(`${lc.row},${toCol}`)) {
 wsDE.row(lc.row).cell(toCol).value(toExcelDateSerial(String(claimData.insurancePeriodTo)));
 }
 break;
 }
 } catch { break; }
 }
 }
 break;
 }

 // Write value one cell to the right of the label — plain values, no style changes
 let cellValue: string | number = DATE_FIELDS.has(key) ? toExcelDateSerial(String(value)) : String(value);
 // Write numeric fields as actual numbers to avoid "stored as text" warning
 if (NUMERIC_FIELDS.has(key)) {
 const num = Number(String(value).replace(/[^\d.]/g, ''));
 if (!isNaN(num) && num > 0) cellValue = num;
 }
 wsDE.row(lc.row).cell(targetCol).value(cellValue);
 console.log(` ${key} → row ${lc.row} col ${targetCol} = ${String(value).substring(0, 60)}`);
 break;
 }
 }
 }

 // ── ASSESSMENT SHEET ──────────────────────────────────────────────────────
 const wsAS = workbook.sheet('Assessment Sheet');
 if (!wsAS) {
 console.warn('Assessment Sheet not found');
 } else if (parts.length > 0) {
 const labelIndex = buildLabelIndex(wsAS, 50, 20);
 const nameOfPartsCell = labelIndex.get(norm('NAME OF PARTS'));

 if (!nameOfPartsCell) {
 console.warn('NAME OF PARTS header not found');
 } else {
 const headerRow = nameOfPartsCell.row;

 // Build column map from the header row
 const colMap: Record<string, number> = {};
 for (let c = 1; c <= 20; c++) {
 const text = safeCellText(wsAS, headerRow, c);
 if (text) colMap[norm(text)] = c;
 }
 console.log('Assessment Sheet columns:', colMap);

 const slNoCol = colMap[norm('SL NO')];
 const nameCol = colMap[norm('NAME OF PARTS')];
 const materialCol = colMap[norm('MATERIAL OF PARTS')];
 const gstCol = colMap[norm('GST (Parts)')];
 const estCostCol = colMap[norm('ESTIMATED COST')];
 const billedCol = colMap[norm('BILLED AMOUNT')];

 // SL1 = headerRow+1, SL2 = headerRow+2 → leave untouched
 // Our parts start at SL3 = headerRow+3
 const startRow = headerRow + 3;

 parts.forEach((part, index) => {
 const r = startRow + index;
 try {
 if (slNoCol) wsAS.row(r).cell(slNoCol).value(index + 3);
 if (nameCol) wsAS.row(r).cell(nameCol).value(part.name);
 if (materialCol) wsAS.row(r).cell(materialCol).value(part.material || 'PLASTIC');
 if (gstCol) wsAS.row(r).cell(gstCol).value(part.gst || '18.00%');
 if (estCostCol) wsAS.row(r).cell(estCostCol).value(parseFloat(part.estimatedCost) || 0);
 if (billedCol) wsAS.row(r).cell(billedCol).value(parseFloat(part.billedAmount) || 0);
 } catch (e) {
 console.error(`Part row ${index} error:`, e);
 }
 });

 console.log(`Wrote ${parts.length} parts starting at row ${startRow}`);
 }
 }

 // Output as buffer — only the cells we touched are different from the template
 const buffer = await workbook.outputAsync();
 return buffer;
}