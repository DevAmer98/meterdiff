import { NextResponse } from "next/server";
import * as XLSX from "xlsx";

function normalizeHeaderKey(h: any) {
  return String(h ?? "")
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "") // strip diacritics
    .replace(/[^0-9a-zA-Z]/g, "") // remove punctuation & spaces
    .trim()
    .toLowerCase();
}

function findKeyByNormalizedTokens(headers: string[], tokens: string[]) {
  const normalizedList = headers.map((h) => ({ orig: h, norm: normalizeHeaderKey(h) }));
  
  console.log("Searching in headers:", normalizedList);
  console.log("Looking for tokens:", tokens);
  
  // exact token match
  for (const t of tokens) {
    const found = normalizedList.find((h) => h.norm === t);
    if (found) {
      console.log(`Found exact match: ${found.orig} (normalized: ${found.norm}) for token: ${t}`);
      return found.orig;
    }
  }
  
  // substring match
  for (const t of tokens) {
    const found = normalizedList.find((h) => h.norm.includes(t));
    if (found) {
      console.log(`Found substring match: ${found.orig} (normalized: ${found.norm}) for token: ${t}`);
      return found.orig;
    }
  }
  
  return null;
}

// --- Convert workbook to JSON with proper header detection ---
function workbookToJsonFromBuffer(buf: ArrayBuffer) {
  const wb = XLSX.read(new Uint8Array(buf), { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];

  // First, get the raw data to inspect the structure
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
  const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

  console.log("Raw data first few rows:", rawData.slice(0, 3));

  if (rawData.length < 1) {
    return [];
  }

  let headers: string[];
  let dataRows: any[][];

  // Check multiple rows to find the best header row
  let headerRowIndex = 0;
  
  for (let i = 0; i < Math.min(3, rawData.length); i++) {
    const row = rawData[i] as any[];
    const nonNullCount = row.filter(cell => cell !== null && cell !== undefined && cell !== '').length;
    const hasRealHeaders = row.some((cell) => {
      const cellStr = String(cell || '').toLowerCase();
      return cellStr.includes('meter') || cellStr.includes('date') || cellStr.includes('energy') || 
             cellStr.includes('usage') || cellStr.includes('point') || cellStr.includes('asset');
    });
    
    // Check if this row contains generic placeholders
    const hasGenericHeaders = row.some((cell) => {
      const cellStr = String(cell || '').toLowerCase();
      return cellStr === 'device' || cellStr === 'column' || cellStr === 'field' || 
             cellStr === 'daily profile';
    });
    
    console.log(`Row ${i}: nonNull=${nonNullCount}, hasRealHeaders=${hasRealHeaders}, hasGeneric=${hasGenericHeaders}`);
    
    // Use this row as headers if it has real headers and good coverage
    if (hasRealHeaders && nonNullCount > 2 && !hasGenericHeaders) {
      headerRowIndex = i;
      console.log(`Selected row ${i} as header row`);
      break;
    }
  }

  headers = (rawData[headerRowIndex] as any[]).map(String);
  dataRows = rawData.slice(headerRowIndex + 1);
  
  console.log(`Using row ${headerRowIndex} as headers`);
  console.log("Final headers:", headers);
  console.log("Data rows count:", dataRows.length);

  // Convert data rows to objects
  const result = dataRows.map(row => {
    const obj: Record<string, any> = {};
    headers.forEach((header, index) => {
      obj[header] = row[index] !== undefined ? row[index] : null;
    });
    return obj;
  });

  // Filter out completely empty rows
  return result.filter(row =>
    Object.values(row).some(v => v !== null && v !== "" && v !== undefined)
  );
}

export async function POST(req: Request) {
  try {
    const form = await req.formData();
    const f1 = form.get("file1") as File | null;
    const f2 = form.get("file2") as File | null;
    const usageKeyOverride = form.get("usageKey") as string | null;
    const joinKeyOverride = form.get("joinKey") as string | null;

    if (!f1 || !f2) {
      return NextResponse.json({ error: "Both files are required (file1 and file2)." }, { status: 400 });
    }

    console.log("Processing files:", f1.name, "and", f2.name);

    const [buf1, buf2] = await Promise.all([f1.arrayBuffer(), f2.arrayBuffer()]);
    const readings = workbookToJsonFromBuffer(buf1);
    const mapping = workbookToJsonFromBuffer(buf2);

    console.log("Readings count:", readings.length);
    console.log("Mapping count:", mapping.length);

    if (!Array.isArray(readings) || readings.length === 0) {
      return NextResponse.json({ error: "Readings file is empty or could not be parsed." }, { status: 400 });
    }

    if (!Array.isArray(mapping) || mapping.length === 0) {
      return NextResponse.json({ error: "Mapping file is empty or could not be parsed." }, { status: 400 });
    }

    const sampleMapRow = mapping[0] || {};
    const mappingHeaders = Object.keys(sampleMapRow);
    const readingsSample = readings[0] || {};
    const readingsHeaders = Object.keys(readingsSample);

    console.log("Mapping headers:", mappingHeaders);
    console.log("Readings headers:", readingsHeaders);
    console.log("Sample mapping row:", sampleMapRow);
    console.log("Sample readings row:", readingsSample);

    // --- Determine usage key ---
    let usageKey: string | null = usageKeyOverride || null;
    
    if (!usageKey) {
      // Look for Usage Point No. in headers
      const usageTokens = [
        "usagepointno", "usagepoint", "usage", "point", "location", "usageno"
      ];
      
      usageKey = findKeyByNormalizedTokens(mappingHeaders, usageTokens);
    }

    if (!usageKey) {
      return NextResponse.json({ 
        error: "Could not find Usage Point column in mapping file.",
        detectedHeaders: mappingHeaders,
        sampleRow: sampleMapRow,
        suggestion: "The mapping file should contain a 'Usage Point No.' column or similar. You can specify it manually using the form field."
      }, { status: 400 });
    }

    console.log("Using usage key:", usageKey);

    // --- Determine join (meter) key in mapping file ---
    let mapMeterKey: string | null = joinKeyOverride || null;
    
    if (!mapMeterKey) {
      // Try to find meter-related columns
      const meterTokens = [
        "meterno", "masterno", "meternumber", "meterid", "meterserial", 
        "serialnumber", "serialno", "serial", "id", "identifier", "meter"
      ];
      
      mapMeterKey = findKeyByNormalizedTokens(mappingHeaders, meterTokens);
    }

    // Fallback to first column that might contain IDs
    if (!mapMeterKey) {
      mapMeterKey = mappingHeaders.find(h => 
        normalizeHeaderKey(h).includes("no") || 
        normalizeHeaderKey(h).includes("id") ||
        normalizeHeaderKey(h).includes("number")
      ) || mappingHeaders[0];
    }

    if (!mapMeterKey) {
      return NextResponse.json({ 
        error: "Could not determine meter join column in mapping file.",
        detectedHeaders: mappingHeaders,
        sampleRow: sampleMapRow
      }, { status: 400 });
    }

    console.log("Using mapping meter key:", mapMeterKey);

    // --- Determine the meter key in readings file ---
    let readingsMeterKey: string | null = null;
    const readingsMeterTokens = [
      "meterno", "masterno", "meterid", "serial", "id", "identifier", "meter"
    ];
    
    readingsMeterKey = findKeyByNormalizedTokens(readingsHeaders, readingsMeterTokens);

    // Fallback strategies
    if (!readingsMeterKey) {
      // Look for columns with "no" or "id" in the name
      readingsMeterKey = readingsHeaders.find(h => 
        normalizeHeaderKey(h).includes("no") || 
        normalizeHeaderKey(h).includes("id") ||
        normalizeHeaderKey(h).includes("number")
      );
    }

    // Last resort - use first column
    if (!readingsMeterKey) {
      readingsMeterKey = readingsHeaders[0];
    }

    if (!readingsMeterKey) {
      return NextResponse.json({ 
        error: "Could not determine meter column in readings file.",
        detectedHeaders: readingsHeaders,
        sampleRow: readingsSample
      }, { status: 400 });
    }

    console.log("Using readings meter key:", readingsMeterKey);

    // --- Build lookup map ---
    const map = new Map<string, string>();
    let mappingProcessed = 0;
    let mappingSkipped = 0;

    for (const row of mapping) {
      const rawVal = row[mapMeterKey];
      const usageVal = row[usageKey];
      
      if (!rawVal || !usageVal) {
        mappingSkipped++;
        continue;
      }
      
      // Normalize the meter key for consistent matching
      const normalizedKey = String(rawVal).trim().toLowerCase();
      const normalizedUsage = String(usageVal).trim();
      
      map.set(normalizedKey, normalizedUsage);
      mappingProcessed++;
    }

    console.log(`Mapping processed: ${mappingProcessed}, skipped: ${mappingSkipped}`);
    console.log("Sample mapping entries:", Array.from(map.entries()).slice(0, 5));

    // --- Merge readings with Usage Point No. ---
    const USAGE_OUTPUT_HEADER = "Usage Point No.";
    let foundCount = 0;
    let notFoundCount = 0;

    const merged = readings.map(r => {
      const val = r[readingsMeterKey!];
      const normalizedVal = val ? String(val).trim().toLowerCase() : "";
      const usage = normalizedVal && map.has(normalizedVal) ? map.get(normalizedVal)! : "NOT FOUND";
      
      if (usage === "NOT FOUND") {
        notFoundCount++;
      } else {
        foundCount++;
      }
      
      // Create a new object with ALL original readings data plus the new Usage Point column
      const mergedRow = { ...r };
      mergedRow[USAGE_OUTPUT_HEADER] = usage;
      return mergedRow;
    });

    console.log(`Merge results: ${foundCount} found, ${notFoundCount} not found`);

    // Show some examples of what we're trying to match
    const readingsSample5 = readings.slice(0, 5).map(r => r[readingsMeterKey!]);
    const mappingKeys5 = Array.from(map.keys()).slice(0, 5);
    
    console.log("Sample readings meter values:", readingsSample5);
    console.log("Sample mapping keys:", mappingKeys5);

    // --- Output Excel ---
    const outWb = XLSX.utils.book_new();
    const outSheet = XLSX.utils.json_to_sheet(merged);
    XLSX.utils.book_append_sheet(outWb, outSheet, "merged");
    
    const outArray = XLSX.write(outWb, { type: "array", bookType: "xlsx" }) as Uint8Array;

    const headers = new Headers();
    headers.set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    headers.set("Content-Disposition", `attachment; filename="meter_merge.xlsx"`);

    return new Response(outArray, { status: 200, headers });
  } catch (err: any) {
    console.error("merge error:", err);
    return NextResponse.json({ 
      error: err?.message ?? "Internal error",
      stack: err?.stack
    }, { status: 500 });
  }
}