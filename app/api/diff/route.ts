import { NextResponse } from "next/server";
import * as XLSX from "xlsx";

export const runtime = "nodejs";

type Row = Record<string, unknown>;

function normKey(k: string) {
  return String(k ?? "").trim().toLowerCase();
}
function squish(k: string) {
  return normKey(k).replace(/[^a-z0-9]+/g, "");
}

const METER_ALIASES = new Set([
  "meter_id","meter","meter id","meter no","meter no.","meterno",
  "meter number","meter number.","meternumber"
]);
const METER_SQUISHED = new Set(["meterid","meter","meterno","meternumber"]);

const VALUE_ALIASES = new Set([
  "value","reading","amount","kwh",
  "active energy import (+a)","active energy import (a)","active energy import"
]);
const VALUE_SQUISHED = new Set([
  "value","reading","amount","kwh","activeenergyimporta","activeenergyimport"
]);

const DATE_ALIASES = new Set([
  "date","reading_date","measurement_date","timestamp","created_date",
  "report_date","تاريخ","التاريخ","datetime","reading date"
]);
const DATE_SQUISHED = new Set([
  "date","readingdate","measurementdate","timestamp","createddate",
  "reportdate","تاريخ","التاريخ","datetime"
]);

const USAGE_POINT_ALIASES = new Set([
  "usage_point_no","usage point no","usage point no.","usagepointno",
  "usage_point","usage point","usagepoint","location","site","address"
]);
const USAGE_POINT_SQUISHED = new Set([
  "usagepointno","usagepoint","location","site","address"
]);

function looksLikeMeter(s: string) {
  const n = normKey(s);
  const q = squish(s);
  return (
    METER_ALIASES.has(n) ||
    METER_SQUISHED.has(q) ||
    q.includes("meter")
  );
}

function looksLikeValue(s: string) {
  const n = normKey(s);
  const q = squish(s);
  return (
    VALUE_ALIASES.has(n) ||
    VALUE_SQUISHED.has(q) ||
    /active\s*energy\s*import/i.test(s) ||
    q === "kwh" ||
    q.includes("value")
  );
}

function looksLikeUsagePoint(s: string) {
  const n = normKey(s);
  const q = squish(s);
  return (
    USAGE_POINT_ALIASES.has(n) ||
    USAGE_POINT_SQUISHED.has(q) ||
    q.includes("usage") ||
    q.includes("location")
  );
}
function looksLikeDate(s: string) {
  const n = normKey(s);
  const q = squish(s);
  return (
    DATE_ALIASES.has(n) ||
    DATE_SQUISHED.has(q) ||
    q.includes("date") ||
    q.includes("time")
  );
}

function parseNumberLike(n: unknown): number | null {
  if (typeof n === "number") return Number.isFinite(n) ? n : null;
  if (n == null) return null;
  let s = String(n).trim();
  if (!s) return null;
  s = s.replace(/\s+/g, "");
  if (s.includes(",") && !s.includes(".")) s = s.replace(/,/g, ".");
  s = s.replace(/[^0-9.-]/g, "");
  const v = Number(s);
  return Number.isFinite(v) ? v : null;
}

function parseDateLike(n: unknown): Date | null {
  if (n == null) return null;
  
  // Handle Excel date serial numbers
  if (typeof n === "number") {
    if (n > 25000 && n < 50000) { // Reasonable range for Excel dates
      try {
        const date = XLSX.SSF.parse_date_code(n);
        return new Date(date.y, date.m - 1, date.d);
      } catch {
        return null;
      }
    }
  }
  
  // Handle date strings
  const s = String(n).trim();
  if (!s) return null;
  
  // Try parsing as Date
  const parsed = new Date(s);
  if (!isNaN(parsed.getTime()) && s.length > 6) {
    return parsed;
  }
  
  // Try common date formats manually
  const datePatterns = [
    /(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/,  // DD/MM/YYYY or DD-MM-YYYY
    /(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/,  // YYYY/MM/DD or YYYY-MM-DD
    /(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2})/,  // DD/MM/YY or DD-MM-YY
  ];
  
  for (const pattern of datePatterns) {
    const match = s.match(pattern);
    if (match) {
      try {
        let day, month, year;
        if (match[3] && match[3].length === 4) { // DD/MM/YYYY
          day = parseInt(match[1]);
          month = parseInt(match[2]);
          year = parseInt(match[3]);
        } else if (match[1].length === 4) { // YYYY/MM/DD
          year = parseInt(match[1]);
          month = parseInt(match[2]);
          day = parseInt(match[3]);
        } else { // DD/MM/YY
          day = parseInt(match[1]);
          month = parseInt(match[2]);
          year = 2000 + parseInt(match[3]);
        }
        
        if (day >= 1 && day <= 31 && month >= 1 && month <= 12) {
          return new Date(year, month - 1, day);
        }
      } catch {
        continue;
      }
    }
  }
  
  return null;
}

function detectColsFromObjectRows(rows: Row[]) {
  if (!rows.length) return null;
  const keys = Object.keys(rows[0]);
  const usable = keys.filter((k) => !k.startsWith("__empty"));
  if (!usable.length) return null;

  const meterKey = usable.find(looksLikeMeter);
  const valueKey = usable.find(looksLikeValue);
  const dateKey = usable.find(looksLikeDate);
  const usagePointKey = usable.find(looksLikeUsagePoint);

  if (!meterKey || !valueKey) return null;
  return { meterKey, valueKey, dateKey, usagePointKey };
}

function detectColsByScanning(ws: XLSX.WorkSheet) {
  const grid: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const MAX_SCAN_ROWS = Math.min(grid.length, 50);

  let headerRowIdx = -1;
  let meterCol = -1;
  let valueCol = -1;
  let dateCol = -1;
  let usagePointCol = -1;

  for (let r = 0; r < MAX_SCAN_ROWS; r++) {
    const row = grid[r] || [];
    for (let c = 0; c < row.length; c++) {
      const cell = row[c];
      if (cell == null) continue;
      const s = String(cell);
      if (meterCol === -1 && looksLikeMeter(s)) meterCol = c;
      if (valueCol === -1 && looksLikeValue(s)) valueCol = c;
      if (dateCol === -1 && looksLikeDate(s)) dateCol = c;
      if (usagePointCol === -1 && looksLikeUsagePoint(s)) usagePointCol = c;
    }
    if (meterCol !== -1 && valueCol !== -1) {
      headerRowIdx = r;
      break;
    }
  }

  if (headerRowIdx === -1) return null;

  return {
    toObjectRows(): Row[] {
      const out: Row[] = [];
      for (let r = headerRowIdx + 1; r < grid.length; r++) {
        const row = grid[r] || [];
        const meter = row[meterCol];
        const value = row[valueCol];
        const date = dateCol !== -1 ? row[dateCol] : null;
        const usagePoint = usagePointCol !== -1 ? row[usagePointCol] : null;
        const obj: Row = { __meter: meter, __value: value, __date: date, __usagePoint: usagePoint };
        out.push(obj);
      }
      return out;
    },
    meterProp: "__meter",
    valueProp: "__value",
    dateProp: "__date",
    usagePointProp: "__usagePoint",
  };
}

function fileToMeterMapWithDatesAndUsage(buf: Buffer) {
  const wb = XLSX.read(buf, { type: "buffer" });
  const sheetName = wb.SheetNames[0];
  if (!sheetName) return { map: new Map<string, number>(), dates: [], usagePoints: new Map<string, string>() };
  const ws = wb.Sheets[sheetName];

  const objectRows: Row[] = XLSX.utils.sheet_to_json(ws, { defval: null });
  let meterKey: string | null = null;
  let valueKey: string | null = null;
  let dateKey: string | null = null;
  let usagePointKey: string | null = null;
  let rowsToUse: Row[] | null = null;

  const foundA = detectColsFromObjectRows(objectRows);
  if (foundA) {
    meterKey = foundA.meterKey;
    valueKey = foundA.valueKey;
    dateKey = foundA.dateKey || null;
    usagePointKey = foundA.usagePointKey || null;
    rowsToUse = objectRows;
  } else {
    const scanned = detectColsByScanning(ws);
    if (!scanned) return { map: new Map<string, number>(), dates: [], usagePoints: new Map<string, string>() };
    meterKey = scanned.meterProp;
    valueKey = scanned.valueProp;
    dateKey = scanned.dateProp;
    usagePointKey = scanned.usagePointProp;
    rowsToUse = scanned.toObjectRows();
  }

  const map = new Map<string, number>();
  const dates: Date[] = [];
  const usagePoints = new Map<string, string>();
  
  for (const r of rowsToUse!) {
    const meter = String(r[meterKey!] ?? "").trim();
    if (!meter) continue;
    const val = parseNumberLike(r[valueKey!]);
    if (val == null) continue;
    
    map.set(meter, (map.get(meter) ?? 0) + val);
    
    // Extract date if available
    if (dateKey && r[dateKey]) {
      const date = parseDateLike(r[dateKey]);
      if (date) {
        dates.push(date);
      }
    }
    
    // Extract usage point if available
    if (usagePointKey && r[usagePointKey]) {
      const usagePoint = String(r[usagePointKey] ?? "").trim();
      if (usagePoint) {
        usagePoints.set(meter, usagePoint);
      }
    }
  }
  
  return { map, dates, usagePoints };
}

function formatDate(date: Date): string {
  return date.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "2-digit", 
    year: "numeric"
  });
}

function getDateRange(dates: Date[]): string {
  if (dates.length === 0) return "";
  
  dates.sort((a, b) => a.getTime() - b.getTime());
  const minDate = dates[0];
  const maxDate = dates[dates.length - 1];
  
  if (minDate.getTime() === maxDate.getTime()) {
    return formatDate(minDate);
  }
  
  return `${formatDate(minDate)} to ${formatDate(maxDate)}`;
}

export async function POST(req: Request) {
  try {
    const form = await req.formData();
    const f1 = form.get("file1");
    const f2 = form.get("file2");
    if (!(f1 instanceof File) || !(f2 instanceof File)) {
      return NextResponse.json({ error: "Please upload file1 and file2." }, { status: 400 });
    }

    const [ab1, ab2] = await Promise.all([f1.arrayBuffer(), f2.arrayBuffer()]);
    const result1 = fileToMeterMapWithDatesAndUsage(Buffer.from(ab1));
    const result2 = fileToMeterMapWithDatesAndUsage(Buffer.from(ab2));
    
    const map1 = result1.map;
    const map2 = result2.map;
    const dates1 = result1.dates;
    const dates2 = result2.dates;
    const usagePoints1 = result1.usagePoints;
    const usagePoints2 = result2.usagePoints;

    const meters = new Set<string>([...map1.keys(), ...map2.keys()]);
    const header = ["meter_id", "usage_point_no", "value_file1", "value_file2", "diff_file2_minus_file1"];

    // Build date ranges (we'll insert above the header in the sheet)
    const dateRange1 = getDateRange(dates1);
    const dateRange2 = getDateRange(dates2);

    const out: (string | number)[][] = [];

    // Add consolidated date range info above the header if available
    if (dateRange1 || dateRange2) {
      let rangeDisplay = "";
      if (dateRange1 && dateRange2) {
        // Human-readable progression shown in-sheet
        rangeDisplay = `${dateRange1} to ${dateRange2}`;
      } else if (dateRange1) {
        rangeDisplay = `${dateRange1} (File 1 only)`;
      } else if (dateRange2) {
        rangeDisplay = `${dateRange2} (File 2 only)`;
      }

      // Put the date range above the headers
      out.push(["Date Range:", rangeDisplay]);
      out.push([""]); // spacer row
    }

    // Now add the header row after the date range
    out.push(header);

    // Add data rows
    for (const m of Array.from(meters).sort()) {
      const v1 = +(map1.get(m) ?? 0);
      const v2 = +(map2.get(m) ?? 0);
      // Get usage point from either file (prefer file2 if both have it)
      const usagePoint = usagePoints2.get(m) || usagePoints1.get(m) || "";
      out.push([m, usagePoint, v1, v2, v2 - v1]);
    }

    const ws = XLSX.utils.aoa_to_sheet(out);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Diff");
    const bin = XLSX.write(wb, { bookType: "xlsx", type: "buffer" });

    // Build consolidatedRange but use ASCII arrow (->) to avoid non-ASCII in header
    const consolidatedRange = dateRange1 && dateRange2 ? `${dateRange1} -> ${dateRange2}` : dateRange1 || dateRange2 || "";

    // Sanitize header: remove non-ASCII bytes and truncate to a safe length (e.g., 200 chars)
    const safeDateRangeHeader = consolidatedRange
      .replace(/[^\x20-\x7E]/g, '') // remove non-ASCII characters
      .slice(0, 200);

    // Build response headers and only include X-Date-Range if non-empty
    const respHeaders: Record<string, string> = {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": 'attachment; filename="meter_diff.xlsx"',
      "Cache-Control": "no-store",
    };
    if (safeDateRangeHeader) {
      respHeaders["X-Date-Range"] = safeDateRangeHeader;
    }

    return new NextResponse(bin, {
      status: 200,
      headers: respHeaders,
    });
  } catch (e: unknown) {
    let msg = "Unknown error";
    if (e instanceof Error) msg = e.message;
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}
