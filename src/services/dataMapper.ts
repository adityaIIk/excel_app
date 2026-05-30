import { ClaimData, ExtractionResult, PartEntry } from '../types/index';

/**
 * Merges multiple ExtractionResult objects into a single one.
 * Non-null values from later results override earlier ones for ClaimData.
 * Parts arrays are concatenated (deduplication by name).
 */
export function mergeResults(results: ExtractionResult[]): ExtractionResult {
 const merged: ExtractionResult = { claimData: {}, parts: [] };
 const seenParts = new Set<string>();

 for (const result of results) {
 // Merge claimData — non-null/empty values win
 for (const [key, value] of Object.entries(result.claimData)) {
 if (value !== null && value !== undefined && value !== '') {
 (merged.claimData as Record<string, unknown>)[key] = value;
 }
 }

 // Deduplicate parts by name
 for (const part of result.parts) {
 const normName = part.name.toLowerCase().replace(/\s+/g, '');
 if (!seenParts.has(normName)) {
 seenParts.add(normName);
 merged.parts.push(part);
 }
 }
 }

 return merged;
}

/**
 * Applies any final post-processing / normalisation to the merged ClaimData
 * before it goes into the Excel sheet.
 */
export function normaliseClaimData(data: ClaimData): ClaimData {
 const out = { ...data };

 // Remove masked/redacted values (e.g. xxxxxxxxxxxx@gmail.com from documents)
 for (const key of Object.keys(out) as (keyof ClaimData)[]) {
 const val = out[key];
 if (typeof val === 'string' && /x{4,}/i.test(val)) {
delete out[key];
 }
 }

 // Ensure insured name has "Mr." prefix (the Excel template pre-fills "Mr.")
 if (out.insuredName &&!out.insuredName.startsWith('Mr.')) {
 out.insuredName = `Mr. ${out.insuredName}`;
 }

 // Normalise vehicle type to match the POLICY EXCESS CALCULATION lookup table
 // Valid values: PRIVATE CAR, TAXI, AMBULANCE, GOODS CARRYING VEHICLE,
 // PASSENGER CARRYING VEHICLE, TWO WHEELER
 if (out.vehicleType) {
 const vt = out.vehicleType.toUpperCase().trim();
 if (
 vt === '2W' ||
 vt.includes('MCWG') ||
 vt.includes('MOTORCYCLE') ||
 vt.includes('2WN') ||
 vt.includes('SCOOTER') ||
 vt.includes('MOPED')
 ) {
 out.vehicleType = 'TWO WHEELER';
 } else if (
 vt === 'LMV' ||
 vt.includes('PRIVATE') ||
 vt.includes('CAR') ||
 vt.includes('SEDAN') ||
 vt.includes('SUV') ||
 vt.includes('HATCHBACK') ||
 vt.includes('LMV')
 ) {
 out.vehicleType = 'PRIVATE CAR';
 } else if (
 vt.includes('TAXI') ||
 vt.includes('COMMERCIAL')
 ) {
 out.vehicleType = 'TAXI';
 } else if (
 vt.includes('AMBULANCE')
 ) {
 out.vehicleType = 'AMBULANCE';
 } else if (
 vt.includes('GOODS') ||
 vt.includes('GCV') ||
 vt.includes('HMV') ||
 vt.includes('HTV') ||
 vt.includes('TRUCK') ||
 vt.includes('LORRY')
 ) {
 out.vehicleType = 'GOODS CARRYING VEHICLE';
 } else if (
 vt.includes('PASSENGER') ||
 vt.includes('PSV') ||
 vt.includes('BUS')
 ) {
 out.vehicleType = 'PASSENGER CARRYING VEHICLE';
 } else if (
 vt.includes('MOTOR') // generic "motor vehicle" → likely two wheeler in Indian context
 ) {
 out.vehicleType = 'TWO WHEELER';
 }
 // If none match, leave as-is — user can correct manually
 }

 // Normalise IDV to include Rs. prefix
 if (out.idv &&!out.idv.toString().startsWith('Rs')) {
 out.idv = `Rs.${out.idv}`;
 }

 // Wrap cause of loss inside the standard insurance template sentence.
 // If extracted from documents (possibly translated from Hindi), embed it.
 // If nothing was extracted, leave the placeholder blank (template stays intact).
 const extractedCause = out.causeOfLoss ? out.causeOfLoss.trim() : '';
 if (extractedCause) {
 // Convert any remaining 1st-person to 3rd-person
 const thirdPerson = extractedCause
 .replace(/\bI was\b/gi, 'the insured was')
 .replace(/\bI am\b/gi, 'the insured is')
 .replace(/\bI\b/gi, 'the insured')
 .replace(/\bmy\b/gi, 'his')
 .replace(/\bme\b/gi, 'him')
 .replace(/\bmyself\b/gi, 'himself');
 out.causeOfLoss = `As per statement of insured and also mentioned in the claim form. While the insured vehicle was en-route. At the place of accident, ${thirdPerson}. Hence the damages sustained.`;
 } else {
 // No cause extracted — leave the field empty so template stays intact
delete out.causeOfLoss;
 }

 return out;
}

/**
 * Normalises each PartEntry — ensures GST is formatted as "18.00%" etc.
 */
export function normaliseparts(parts: PartEntry[]): PartEntry[] {
 return parts.map((p) => {
 let gst = p.gst ?? '18.00%';
 // If gst is a plain number, format it
 const gstNum = parseFloat(gst.toString().replace('%', ''));
 if (!isNaN(gstNum)) {
 gst = `${gstNum.toFixed(2)}%`;
 }

 return { ...p, gst };
 });
}