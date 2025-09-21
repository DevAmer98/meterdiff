import { NextResponse } from "next/server";
import * as XLSX from "xlsx";

export const runtime = "nodejs";

type Row = Record<string, unknown>; // replace 'any' with unknown

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

function detectColsFromObjectRows(rows: Row[]) {
  if (!rows.length) return null;
  const keys = Object.keys(rows[0]);
  const usable = keys.filter((k) => !k.startsWith("__empty"));
  if (!usable.length) return null;

  const meterKey = usable.find(looksLikeMeter);
  const valueKey = usable.find(looksLikeValue);

  if (!meterKey || !valueKey) return null;
  return { meterKey, valueKey };
}

function detectColsByScanning(ws: XLSX.WorkSheet) {
  const grid: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const MAX_SCAN_ROWS = Math.min(grid.length, 50);

  let headerRowIdx = -1;
  let meterCol = -1;
  let valueCol = -1;

  for (let r = 0; r < MAX_SCAN_ROWS; r++) {
    const row = grid[r] || [];
    for (let c = 0; c < row.length; c++) {
      const cell = row[c];
      if (cell == null) continue;
      const s = String(cell);
      if (meterCol === -1 && looksLikeMeter(s)) meterCol = c;
      if (valueCol === -1 && looksLikeValue(s)) valueCol = c;
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
        const obj: Row = { __meter: meter, __value: value };
        out.push(obj);
      }
      return out;
    },
    meterProp: "__meter",
    valueProp: "__value",
  };
}

function fileToMeterMap(buf: Buffer) {
  const wb = XLSX.read(buf, { type: "buffer" });
  const sheetName = wb.SheetNames[0];
  if (!sheetName) return new Map<string, number>();
  const ws = wb.Sheets[sheetName];

  const objectRows: Row[] = XLSX.utils.sheet_to_json(ws, { defval: null });
  let meterKey: string | null = null;
  let valueKey: string | null = null;
  let rowsToUse: Row[] | null = null;

  const foundA = detectColsFromObjectRows(objectRows);
  if (foundA) {
    meterKey = foundA.meterKey;
    valueKey = foundA.valueKey;
    rowsToUse = objectRows;
  } else {
    const scanned = detectColsByScanning(ws);
    if (!scanned) return new Map<string, number>();
    meterKey = scanned.meterProp;
    valueKey = scanned.valueProp;
    rowsToUse = scanned.toObjectRows();
  }

  const map = new Map<string, number>();
  for (const r of rowsToUse!) {
    const meter = String(r[meterKey!] ?? "").trim();
    if (!meter) continue;
    const val = parseNumberLike(r[valueKey!]);
    if (val == null) continue;
    map.set(meter, (map.get(meter) ?? 0) + val);
  }
  return map;
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
    const map1 = fileToMeterMap(Buffer.from(ab1));
    const map2 = fileToMeterMap(Buffer.from(ab2));

    const meters = new Set<string>([...map1.keys(), ...map2.keys()]);
    const header = ["meter_id", "value_file1", "value_file2", "diff_file2_minus_file1"];
    const out: (string | number)[][] = [header];

    for (const m of Array.from(meters).sort()) {
      const v1 = +(map1.get(m) ?? 0);
      const v2 = +(map2.get(m) ?? 0);
      out.push([m, v1, v2, v2 - v1]);
    }

    const ws = XLSX.utils.aoa_to_sheet(out);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Diff");
    const bin = XLSX.write(wb, { bookType: "xlsx", type: "buffer" });

    return new NextResponse(bin, {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": 'attachment; filename="meter_diff.xlsx"',
        "Cache-Control": "no-store",
      },
    });
  } catch (e: unknown) {
    let msg = "Unknown error";
    if (e instanceof Error) msg = e.message;
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}
