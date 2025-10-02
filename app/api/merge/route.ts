import { NextResponse } from "next/server";
import * as XLSX from "xlsx";

export const runtime = "nodejs";

type SheetRow = Record<string, unknown>;

function normalizeHeaderKey(h: unknown): string {
  return String(h ?? "")
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^0-9a-zA-Z]/g, "")
    .trim()
    .toLowerCase();
}

function findKeyByNormalizedTokens(headers: string[], tokens: string[]): string | null {
  const normalizedList = headers.map((h) => ({ orig: h, norm: normalizeHeaderKey(h) }));
  console.log("Searching in headers:", normalizedList);
  console.log("Looking for tokens (raw):", tokens);

  const normTokens = tokens.map((t) => normalizeHeaderKey(t));

  // exact token match
  for (const t of normTokens) {
    const found = normalizedList.find((h) => h.norm === t);
    if (found) {
      console.log(`Found exact match: ${found.orig} (normalized: ${found.norm}) for token: ${t}`);
      return found.orig;
    }
  }

  // substring match
  for (const t of normTokens) {
    const found = normalizedList.find((h) => h.norm.includes(t));
    if (found) {
      console.log(`Found substring match: ${found.orig} (normalized: ${found.norm}) for token: ${t}`);
      return found.orig;
    }
  }
  return null;
}

function workbookToJsonFromBuffer(buf: ArrayBuffer): SheetRow[] {
  const wb = XLSX.read(new Uint8Array(buf), { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rawData = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null }) as unknown[][];

  console.log("Raw data first few rows:", rawData.slice(0, 3));
  if (rawData.length < 1) return [];

  // pick header row
  let headerRowIndex = 0;
  for (let i = 0; i < Math.min(3, rawData.length); i++) {
    const row = (rawData[i] ?? []) as unknown[];
    const nonNullCount = row.filter((cell) => cell !== null && cell !== undefined && cell !== "").length;

    const hasRealHeaders = row.some((cell) => {
      const s = String(cell ?? "").toLowerCase();
      return s.includes("meter") || s.includes("date") || s.includes("energy") || s.includes("usage") || s.includes("point") || s.includes("asset");
    });

    const hasGenericHeaders = row.some((cell) => {
      const s = String(cell ?? "").toLowerCase();
      return s === "device" || s === "column" || s === "field" || s === "daily profile";
    });

    console.log(`Row ${i}: nonNull=${nonNullCount}, hasRealHeaders=${hasRealHeaders}, hasGeneric=${hasGenericHeaders}`);
    if (hasRealHeaders && nonNullCount > 2 && !hasGenericHeaders) {
      headerRowIndex = i;
      console.log(`Selected row ${i} as header row`);
      break;
    }
  }

  const headers: string[] = (rawData[headerRowIndex] as unknown[]).map((v) => String(v ?? ""));
  const dataRows = rawData.slice(headerRowIndex + 1) as unknown[][];
  console.log(`Using row ${headerRowIndex} as headers`);
  console.log("Final headers:", headers);
  console.log("Data rows count:", dataRows.length);

  const result: SheetRow[] = dataRows.map((row) => {
    const obj: SheetRow = {};
    headers.forEach((header, index) => {
      const val = (row as unknown[])[index];
      obj[header] = val !== undefined ? val : null;
    });
    return obj;
  });

  return result.filter((row) => Object.values(row).some((v) => v !== null && v !== "" && v !== undefined));
}

export async function POST(req: Request) {
  try {
    const form = await req.formData();
    const f1 = form.get("file1") as File | null;
    const f2 = form.get("file2") as File | null;
    const usageKeyOverride = (form.get("usageKey") as string | null) || null;
    const joinKeyOverride = (form.get("joinKey") as string | null) || null;

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

    const sampleMapRow: SheetRow = mapping[0] || {};
    const mappingHeaders = Object.keys(sampleMapRow);
    const readingsSample: SheetRow = readings[0] || {};
    const readingsHeaders = Object.keys(readingsSample);

    console.log("Mapping headers:", mappingHeaders);
    console.log("Readings headers:", readingsHeaders);
    console.log("Sample mapping row:", sampleMapRow);
    console.log("Sample readings row:", readingsSample);

    // Usage key in mapping
    let usageKey: string | null = usageKeyOverride;
    if (!usageKey) {
      const usageTokens = ["usagepointno", "usagepoint", "usage", "point", "location", "usageno"];
      usageKey = findKeyByNormalizedTokens(mappingHeaders, usageTokens);
    }
    if (!usageKey) {
      return NextResponse.json(
        {
          error: "Could not find Usage Point column in mapping file.",
          detectedHeaders: mappingHeaders,
          sampleRow: sampleMapRow,
          suggestion:
            "The mapping file should contain a 'Usage Point No.' column or similar. You can specify it manually using the form field.",
        },
        { status: 400 }
      );
    }
    console.log("Using usage key:", usageKey);

    // Meter key in mapping
    let mapMeterKey: string | null = joinKeyOverride;
    if (!mapMeterKey) {
      const meterTokens = ["meterno", "masterno", "meternumber", "meterid", "meterserial", "serialnumber", "serialno", "serial", "id", "identifier", "meter"];
      mapMeterKey = findKeyByNormalizedTokens(mappingHeaders, meterTokens);
    }
    if (!mapMeterKey) {
      mapMeterKey = mappingHeaders.find((h) => {
        const k = normalizeHeaderKey(h);
        return k.includes("no") || k.includes("id") || k.includes("number");
      }) || mappingHeaders[0];
    }
    if (!mapMeterKey) {
      return NextResponse.json(
        { error: "Could not determine meter join column in mapping file.", detectedHeaders: mappingHeaders, sampleRow: sampleMapRow },
        { status: 400 }
      );
    }
    console.log("Using mapping meter key:", mapMeterKey);

    // Meter key in readings
    let readingsMeterKey: string | null = null;
    readingsMeterKey = findKeyByNormalizedTokens(readingsHeaders, ["meterno", "masterno", "meterid", "serial", "id", "identifier", "meter"]);
    if (!readingsMeterKey) {
      readingsMeterKey = readingsHeaders.find((h) => {
        const k = normalizeHeaderKey(h);
        return k.includes("no") || k.includes("id") || k.includes("number");
      })!;
    }
    if (!readingsMeterKey) readingsMeterKey = readingsHeaders[0];
    if (!readingsMeterKey) {
      return NextResponse.json(
        { error: "Could not determine meter column in readings file.", detectedHeaders: readingsHeaders, sampleRow: readingsSample },
        { status: 400 }
      );
    }
    console.log("Using readings meter key:", readingsMeterKey);

    // Build mapping
    const map = new Map<string, string>();
    let mappingProcessed = 0;
    let mappingSkipped = 0;
    for (const row of mapping) {
      const rawVal = row[mapMeterKey] as unknown;
      const usageVal = row[usageKey] as unknown;
      if (rawVal === null || rawVal === undefined || rawVal === "" || usageVal === null || usageVal === undefined || usageVal === "") {
        mappingSkipped++;
        continue;
      }
      const normalizedKey = String(rawVal).trim().toLowerCase();
      const normalizedUsage = String(usageVal).trim();
      map.set(normalizedKey, normalizedUsage);
      mappingProcessed++;
    }
    console.log(`Mapping processed: ${mappingProcessed}, skipped: ${mappingSkipped}`);
    console.log("Sample mapping entries:", Array.from(map.entries()).slice(0, 5));

    // Merge
    const USAGE_OUTPUT_HEADER = "Usage Point No.";
    let foundCount = 0;
    let notFoundCount = 0;

    const merged: SheetRow[] = readings.map((r) => {
      const val = r[readingsMeterKey as string] as unknown;
      const normalizedVal = val ? String(val).trim().toLowerCase() : "";
      const usage = normalizedVal && map.has(normalizedVal) ? (map.get(normalizedVal) as string) : "NOT FOUND";
      if (usage === "NOT FOUND") notFoundCount++; else foundCount++;
      return { ...r, [USAGE_OUTPUT_HEADER]: usage };
    });

    console.log(`Merge results: ${foundCount} found, ${notFoundCount} not found`);
    console.log("Sample readings meter values:", readings.slice(0, 5).map((r) => r[readingsMeterKey as string]));
    console.log("Sample mapping keys:", Array.from(map.keys()).slice(0, 5));

    // --- Output Excel ---
    const outWb = XLSX.utils.book_new();
    const outSheet = XLSX.utils.json_to_sheet(merged);
    XLSX.utils.book_append_sheet(outWb, outSheet, "merged");

    const outArray = XLSX.write(outWb, { type: "array", bookType: "xlsx" }) as Uint8Array;

    // Make a clean ArrayBuffer (no SharedArrayBuffer typing issues)
    const outAb: ArrayBuffer = outArray.buffer.slice(outArray.byteOffset, outArray.byteOffset + outArray.byteLength);

    const responseHeaders = new Headers();
    responseHeaders.set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    responseHeaders.set("Content-Disposition", `attachment; filename="meter_merge.xlsx"`);
    responseHeaders.set("Cache-Control", "no-store");

    return new Response(outAb, { status: 200, headers: responseHeaders });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : "Internal error";
    const stack = err instanceof Error ? err.stack : undefined;
    console.error("merge error:", err);
    return NextResponse.json({ error: message, stack }, { status: 500 });
  }
}
