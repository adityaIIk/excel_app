import { GoogleGenerativeAI } from '@google/generative-ai';
import Groq from 'groq-sdk';
import { ExtractionResult } from '../types/index';

// ── Provider config ──────────────────────────────────────────────────────────
// Primary: Gemini 2.0 Flash (best free OCR)
// Fallback: Groq LLaMA 4 Maverick (unlimited free, decent OCR)
const GEMINI_MODEL = 'gemini-2.0-flash';
const GROQ_MODEL = 'meta-llama/llama-4-maverick-17b-128e-instruct';
const MAX_RETRIES = 3;

const EXTRACTION_PROMPT = `You are an expert OCR assistant for Indian motor insurance claim documents.

You will receive a document image (driving licence, insurance policy, vehicle RC, claim form, repair estimate, etc.).

Analyse the image and extract every visible field. Return ONLY a single valid JSON object (no markdown, no explanation).

JSON structure (use null for fields not found):
{
 "claimData": {
 "claimNumber": null,
 "policyNumber": null,
 "insuredName": null,
 "insuredPhone": null,
 "insuredEmail": null,
 "vehicleType": null,
 "financerName": null,
 "insurancePeriodFrom": null,
 "insurancePeriodTo": null,
 "idv": null,
 "regNumber": null,
 "regDate": null,
 "vehicleSaleDate": null,
 "chassisNumber": null,
 "engineNumber": null,
 "makeModel": null,
 "colour": null,
 "cubicCapacity": null,
 "seatingCapacity": null,
 "accidentDate": null,
 "accidentPlace": null,
 "dateOfIntimation": null,
 "estimatedLoss": null,
 "causeOfLoss": null,
 "driverName": null,
 "dlNumber": null,
 "dlValidUpto": null,
 "dlIssuedOn": null,
 "dlIssuingAuthority": null,
 "licenseType": null
 },
 "parts": []
}

IMPORTANT RULES:
- If any image is a repair estimate / parts bill: populate "parts" as array of objects:
 { "name": "part name", "material": "PLASTIC|GLASS|METAL|RUBBER", "gst": "18.00%", "estimatedCost": "7087", "billedAmount": "6005.93" }
 - estimatedCost = MRP column; billedAmount = Price/Unit after discount
 - Infer material: lamps/indicators/glass → GLASS, axle/fork/tube/bracket metal parts → METAL, covers/fairings/panels → PLASTIC
 - Include ALL parts rows visible
- licenseType: vehicle class code only (e.g. "MCWG", "LMV", "MCWG/LMV")
- IDV: include "Rs." prefix if present
- causeOfLoss: look for fields labelled "Short Description of Accident/Incident", "Cause of Accident", "Cause of Loss", or similar. If the text is in Hindi or any regional Indian language, translate it accurately to English first, then use the English translation as the value. IMPORTANT: Always write causeOfLoss in THIRD PERSON from a surveyor's perspective (e.g. "the insured was riding his vehicle from his residence..." NOT "I was riding my vehicle..."). Never use first person (I, my, me).
- ALL text values in the JSON must be in English. If any extracted value is in Hindi or another Indian language, translate it to English.
- ALL date values must be in DD/MM/YYYY format (e.g. "15/03/2025"). Convert any other date format to DD/MM/YYYY.
- Return ONLY the JSON object, nothing else`;

/** Parse retry-after delay in ms from a rate-limit error message */
function getRetryDelayMs(err: Error): number {
 const match = err.message.match(/(?:retryDelay|retry_after|Please retry in)['":\s]+(\d+)/i);
 if (match) return (parseInt(match[1], 10) + 2) * 1000;
 return 15000; // default 15s
}

function sleep(ms: number): Promise<void> {
 return new Promise((resolve) => setTimeout(resolve, ms));
}

// ── Gemini extraction (primary) ──────────────────────────────────────────────
async function extractWithGemini(
 genAI: GoogleGenerativeAI,
 file: { buffer: Buffer; mimetype: string },
 index: number,
 total: number
): Promise<{ claimData: Record<string, unknown>; parts: unknown[] }> {
 const model = genAI.getGenerativeModel({ model: GEMINI_MODEL });
 let lastErr: Error | null = null;

 for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
 try {
 console.log(`Image ${index + 1}/${total} → Gemini Flash (attempt ${attempt})...`);
 const result = await model.generateContent([
 { text: EXTRACTION_PROMPT },
 {
 inlineData: {
 mimeType: file.mimetype,
 data: file.buffer.toString('base64'),
 },
 },
 ]);
 const text = result.response.text().trim();
 const jsonText = text.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '');
 const parsed = JSON.parse(jsonText);
 return { claimData: parsed.claimData || {}, parts: parsed.parts || [] };
 } catch (err: unknown) {
 lastErr = err instanceof Error ? err : new Error(String(err));
 const isRateLimit = lastErr.message.includes('429') || lastErr.message.includes('quota');
 if (isRateLimit && attempt < MAX_RETRIES) {
 const delay = getRetryDelayMs(lastErr);
 console.warn(`Gemini rate limit (attempt ${attempt}). Waiting ${delay / 1000}s...`);
 await sleep(delay);
 continue;
 }
 throw lastErr;
 }
 }
 throw lastErr ?? new Error(`Gemini failed on image ${index + 1}`);
}

// ── Groq extraction (fallback) ──────────────────────────────────────────────
async function extractWithGroq(
 groq: Groq,
 file: { buffer: Buffer; mimetype: string },
 index: number,
 total: number
): Promise<{ claimData: Record<string, unknown>; parts: unknown[] }> {
 const base64 = file.buffer.toString('base64');
 let lastErr: Error | null = null;

 for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
 try {
 console.log(`Image ${index + 1}/${total} → Groq Maverick fallback (attempt ${attempt})...`);
 const response = await groq.chat.completions.create({
 model: GROQ_MODEL,
 messages: [
 {
 role: 'user',
 content: [
 { type: 'text', text: EXTRACTION_PROMPT },
 { type: 'image_url', image_url: { url: `data:${file.mimetype};base64,${base64}` } },
 ],
 },
 ],
 max_tokens: 2048,
 temperature: 0,
 });

 const text = (response.choices[0].message.content ?? '').trim();
 const jsonText = text.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '');
 const parsed = JSON.parse(jsonText);
 return { claimData: parsed.claimData || {}, parts: parsed.parts || [] };
 } catch (err: unknown) {
 lastErr = err instanceof Error ? err : new Error(String(err));
 const is429 = lastErr.message.includes('429') || lastErr.message.includes('rate_limit');
 if (is429 && attempt < MAX_RETRIES) {
 const delay = getRetryDelayMs(lastErr);
 console.warn(`Groq rate limit (attempt ${attempt}). Waiting ${delay / 1000}s...`);
 await sleep(delay);
 continue;
 }
 throw lastErr;
 }
 }
 throw lastErr ?? new Error(`Groq failed on image ${index + 1}`);
}

export async function extractFromImages(
 files: Array<{ buffer: Buffer; mimetype: string }>
): Promise<ExtractionResult> {
 const geminiKey = process.env.GEMINI_API_KEY;
 const groqKey = process.env.GROQ_API_KEY;

 if (!geminiKey &&!groqKey) {
 throw new Error('Set GEMINI_API_KEY (primary) or GROQ_API_KEY (fallback) in .env');
 }

 const useGemini =!!geminiKey;
 const genAI = geminiKey ? new GoogleGenerativeAI(geminiKey) : null;
 const groq = groqKey ? new Groq({ apiKey: groqKey }) : null;

 console.log(`Using ${useGemini ? 'Gemini 2.0 Flash (primary)' : 'Groq Maverick (fallback)'}`);

 const merged: ExtractionResult = { claimData: {}, parts: [] };

 for (let i = 0; i < files.length; i++) {
 let result: { claimData: Record<string, unknown>; parts: unknown[] };

 try {
 if (useGemini && genAI) {
 result = await extractWithGemini(genAI, files[i], i, files.length);
 } else if (groq) {
 result = await extractWithGroq(groq, files[i], i, files.length);
 } else {
 throw new Error('No API key available');
 }
 } catch (err) {
 // If Gemini fails, try Groq fallback
 if (useGemini && groq) {
 console.warn(`Gemini failed on image ${i + 1}, falling back to Groq...`);
 result = await extractWithGroq(groq, files[i], i, files.length);
 } else {
 throw err;
 }
 }

 // Merge claimData: first non-null value for each field wins
 for (const [key, value] of Object.entries(result.claimData)) {
 if (value !== null && value !== undefined && value !== '') {
 if (!(merged.claimData as Record<string, unknown>)[key]) {
 (merged.claimData as Record<string, unknown>)[key] = value;
 }
 }
 }

 // Append parts (only repair estimate image will have these)
 if (result.parts.length > 0) {
 merged.parts.push(...(result.parts as typeof merged.parts));
 }

 // Brief pause between images to stay within per-minute limits
 if (i < files.length - 1) await sleep(1500);
 }

 return merged;
}