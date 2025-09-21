"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import { Upload, FileSpreadsheet, X, Loader2, Download, Clock } from "lucide-react";
import clsx from "clsx";

type DropState = "idle" | "over1" | "over2";

function formatDeadlineForRiyadh(d: Date) {
  return d.toLocaleString("en-GB", {
    year: "numeric",
    month: "short",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    timeZone: "Asia/Riyadh",
  });
}

function formatCountdown(ms: number) {
  if (ms <= 0) return "00d 00h 00m 00s";
  const s = Math.floor(ms / 1000);
  const days = Math.floor(s / (60 * 60 * 24));
  const hours = Math.floor((s % (60 * 60 * 24)) / (60 * 60));
  const minutes = Math.floor((s % (60 * 60)) / 60);
  const seconds = s % 60;
  return `${String(days).padStart(2, "0")}d ${String(hours).padStart(2, "0")}h ${String(minutes).padStart(2, "0")}m ${String(seconds).padStart(2, "0")}s`;
}

export default function Home() {
  const [file1, setFile1] = useState<File | null>(null);
  const [file2, setFile2] = useState<File | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [dropState, setDropState] = useState<DropState>("idle");

  // --- Deadline: one year from now ---
  const initialNow = useMemo(() => new Date(), []);
  const [now, setNow] = useState<Date>(initialNow);

  const deadline = useMemo(() => {
    const d = new Date(initialNow);
    d.setFullYear(d.getFullYear() + 1);
    return d;
  }, [initialNow]);

  useEffect(() => {
    const id = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(id);
  }, []);

  const msLeft = Math.max(0, deadline.getTime() - now.getTime());
  const isExpired = msLeft <= 0;

  const onFilesPicked = (which: 1 | 2, files: FileList | null) => {
    const f = files?.[0] ?? null;
    if (which === 1) setFile1(f);
    else setFile2(f);
  };

  const onDrop = useCallback(
    (ev: React.DragEvent, which: 1 | 2) => {
      ev.preventDefault();
      setDropState("idle");
      onFilesPicked(which, ev.dataTransfer.files);
    },
    []
  );

  async function onSubmit(e: React.FormEvent) {
    e.preventDefault();
    setError(null);
    setDownloadUrl(null);

    if (isExpired) {
      setError("This tool has reached its deadline and is no longer available.");
      return;
    }

    if (!file1 || !file2) {
      setError("Please select both files.");
      return;
    }

    setBusy(true);
    try {
      const fd = new FormData();
      fd.append("file1", file1);
      fd.append("file2", file2);

      const res = await fetch("meterdiff/api/diff", { method: "POST", body: fd });
      if (!res.ok) {
        let msg = `Request failed (${res.status})`;
        try {
          const data = await res.json();
          if (data?.error) msg = data.error;
        } catch {}
        throw new Error(msg);
      }

      const arr = await res.arrayBuffer();
      const url = URL.createObjectURL(
        new Blob([arr], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        })
      );
      setDownloadUrl(url);
    } catch (err: unknown) {
      if (err instanceof Error) {
        setError(err.message);
      } else {
        setError(String(err));
      }
    } finally {
      setBusy(false);
    }
  }

  return (
    <main className="relative min-h-screen bg-gradient-to-br from-zinc-950 via-zinc-900 to-black text-zinc-100">
      <div className="pointer-events-none absolute inset-0 [mask-image:radial-gradient(50%_50%_at_50%_50%,#000_20%,transparent_70%)]">
        <div className="absolute left-1/2 top-24 h-72 w-72 -translate-x-1/2 rounded-full bg-gradient-to-tr from-indigo-500/20 via-sky-400/20 to-fuchsia-500/20 blur-3xl" />
      </div>

      <div className="container mx-auto flex min-h-screen items-center justify-center p-6">
        <div className="relative w-full max-w-2xl rounded-3xl border border-white/10 bg-white/5 p-8 shadow-2xl backdrop-blur-md overflow-hidden">
          <div
            className="pointer-events-none absolute inset-0 -z-0 flex items-center justify-center"
            style={{
              backgroundImage: `url('meterdiff/image.png')`,
              backgroundRepeat: "no-repeat",
              backgroundPosition: "center",
              backgroundSize: "contain",
              opacity: 0.1,
              filter: "blur(1px) saturate(0.9)",
            }}
          />

          <div className="relative z-10">
            <header className="mb-6 flex items-start justify-between gap-4">
              <div>
                <h1 className="text-3xl font-semibold tracking-tight">Meter Diff</h1>
                <p className="mt-2 text-sm text-zinc-400">
                  Upload two Excel files containing <code>meter_id</code> and <code>value</code>. We’ll
                  aggregate by meter and generate a results workbook.
                </p>
              </div>

              <div className="ml-4 flex flex-col items-end">
                <div
                  className={clsx(
                    "inline-flex items-center gap-2 rounded-full px-3 py-1 text-xs font-medium",
                    isExpired ? "bg-red-700/30 text-red-200 border border-red-700/40" : "bg-emerald-700/10 text-emerald-200 border border-emerald-400/20"
                  )}
                  title="Deadline"
                >
                  <Clock className="h-4 w-4" />
                  {isExpired ? "Expired" : "Deadline"}
                </div>

                <div className="mt-2 text-right text-xs text-zinc-300">
                  <div className="text-xs">Ends on</div>
                  <div className="text-sm font-medium">{formatDeadlineForRiyadh(deadline)}</div>
                  <div className={clsx("mt-1 text-xs", isExpired ? "text-red-300" : "text-zinc-400")}>
                    {isExpired ? "Deadline passed" : `Time left: ${formatCountdown(msLeft)}`}
                  </div>
                </div>
              </div>
            </header>

            <form onSubmit={onSubmit} className="space-y-6">
              {/* file 1 */}
              <div>
                <label className="mb-2 block text-sm font-medium text-zinc-300">File 1</label>
                <div
                  onDragOver={(e) => {
                    e.preventDefault();
                    setDropState("over1");
                  }}
                  onDragLeave={() => setDropState("idle")}
                  onDrop={(e) => onDrop(e, 1)}
                  className={clsx(
                    "rounded-2xl border-2 border-dashed p-6 transition-colors",
                    dropState === "over1" ? "border-sky-400/70 bg-sky-400/5" : "border-white/10 hover:border-white/20"
                  )}
                >
                  <div className="flex items-center gap-4">
                    <div className="rounded-xl bg-white/5 p-3">
                      <Upload className="h-5 w-5 text-zinc-300" />
                    </div>
                    <div className="flex-1">
                      <p className="text-sm text-zinc-300">
                        Drag & drop Excel here, or{" "}
                        <label className="cursor-pointer underline decoration-dotted underline-offset-4">
                          <span>
                            <input
                              type="file"
                              accept=".xlsx,.xls"
                              className="hidden"
                              onChange={(e) => onFilesPicked(1, e.target.files)}
                            />
                            browse
                          </span>
                        </label>
                      </p>
                      <p className="text-xs text-zinc-500">Accepted: .xlsx, .xls</p>
                    </div>
                  </div>

                  {file1 && (
                    <div className="mt-4 flex items-center justify-between rounded-xl border border-white/10 bg-white/5 px-3 py-2">
                      <div className="flex items-center gap-2">
                        <FileSpreadsheet className="h-4 w-4" />
                        <span className="text-sm">
                          {file1.name} <span className="text-zinc-400">({(file1.size / 1024).toFixed(1)} KB)</span>
                        </span>
                      </div>
                      <button
                        type="button"
                        onClick={() => setFile1(null)}
                        className="rounded-lg p-1 hover:bg-white/10"
                        aria-label="Remove file 1"
                      >
                        <X className="h-4 w-4" />
                      </button>
                    </div>
                  )}
                </div>
              </div>

              {/* file 2 */}
              <div>
                <label className="mb-2 block text-sm font-medium text-zinc-300">File 2</label>
                <div
                  onDragOver={(e) => {
                    e.preventDefault();
                    setDropState("over2");
                  }}
                  onDragLeave={() => setDropState("idle")}
                  onDrop={(e) => onDrop(e, 2)}
                  className={clsx(
                    "rounded-2xl border-2 border-dashed p-6 transition-colors",
                    dropState === "over2" ? "border-fuchsia-400/70 bg-fuchsia-400/5" : "border-white/10 hover:border-white/20"
                  )}
                >
                  <div className="flex items-center gap-4">
                    <div className="rounded-xl bg-white/5 p-3">
                      <Upload className="h-5 w-5 text-zinc-300" />
                    </div>
                    <div className="flex-1">
                      <p className="text-sm text-zinc-300">
                        Drag & drop Excel here, or{" "}
                        <label className="cursor-pointer underline decoration-dotted underline-offset-4">
                          <span>
                            <input
                              type="file"
                              accept=".xlsx,.xls"
                              className="hidden"
                              onChange={(e) => onFilesPicked(2, e.target.files)}
                            />
                            browse
                          </span>
                        </label>
                      </p>
                      <p className="text-xs text-zinc-500">Accepted: .xlsx, .xls</p>
                    </div>
                  </div>

                  {file2 && (
                    <div className="mt-4 flex items-center justify-between rounded-xl border border-white/10 bg-white/5 px-3 py-2">
                      <div className="flex items-center gap-2">
                        <FileSpreadsheet className="h-4 w-4" />
                        <span className="text-sm">
                          {file2.name} <span className="text-zinc-400">({(file2.size / 1024).toFixed(1)} KB)</span>
                        </span>
                      </div>
                      <button
                        type="button"
                        onClick={() => setFile2(null)}
                        className="rounded-lg p-1 hover:bg-white/10"
                        aria-label="Remove file 2"
                      >
                        <X className="h-4 w-4" />
                      </button>
                    </div>
                  )}
                </div>
              </div>

              {/* actions */}
              <div className="flex items-center gap-3 pt-2">
                <button
                  type="submit"
                  disabled={busy || !file1 || !file2 || isExpired}
                  className={clsx(
                    "inline-flex items-center gap-2 rounded-xl px-4 py-2 font-medium transition-colors",
                    busy || !file1 || !file2 || isExpired
                      ? "bg-white/10 text-zinc-300 cursor-not-allowed"
                      : "bg-white text-black hover:bg-zinc-200"
                  )}
                >
                  {busy ? <Loader2 className="h-4 w-4 animate-spin" /> : <Upload className="h-4 w-4" />}
                  {busy ? "Processing..." : isExpired ? "Deadline passed" : "Compute & Download"}
                </button>

                {(file1 || file2) && (
                  <button
                    type="button"
                    onClick={() => {
                      setFile1(null);
                      setFile2(null);
                      setDownloadUrl(null);
                      setError(null);
                    }}
                    className="rounded-xl border border-white/10 px-3 py-2 text-sm text-zinc-300 hover:bg-white/10"
                  >
                    Reset
                  </button>
                )}
              </div>

              {error && (
                <p className="rounded-xl border border-red-900/40 bg-red-900/20 px-3 py-2 text-sm text-red-200">
                  {error}
                </p>
              )}
            </form>

            {downloadUrl && (
              <div className="mt-6">
                <a
                  href={downloadUrl}
                  download="meter_diff.xlsx"
                  className="inline-flex items-center gap-2 rounded-xl border border-emerald-400/40 bg-emerald-400/10 px-4 py-2 text-emerald-200 hover:bg-emerald-400/20"
                >
                  <Download className="h-4 w-4" />
                  Download result (meter_diff.xlsx)
                </a>
              </div>
            )}
          </div>
        </div>
      </div>
    </main>
  );
}
