import { useCallback, useRef, useState } from "react";
import type { CategoryTotals, Config } from "./lib/types";
import { loadConfig } from "./lib/config";
import { processInvoice, type ProcessResult } from "./lib/process";
import { appendToMaster } from "./lib/overlay";
import CategoryEditor from "./CategoryEditor";

interface Download {
  name: string;
  url: string;
}

interface CardResult extends Partial<ProcessResult> {
  name: string;
  status: "busy" | "ok" | "err";
  url?: string;
}

function toDownload(name: string, bytes: Uint8Array): Download {
  const blob = new Blob([bytes as BlobPart], { type: "application/pdf" });
  return { name, url: URL.createObjectURL(blob) };
}

function fmtMoney(v: number): string {
  return (
    "$" +
    v.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })
  );
}

function fmtSize(bytes: number): string {
  return bytes > 1024 * 1024
    ? (bytes / (1024 * 1024)).toFixed(1) + " MB"
    : Math.round(bytes / 1024) + " KB";
}

function catColor(config: Config, label: string): string {
  const code = label.split(" ")[0];
  const cat = config.categories.find((c) => c.code === code);
  return cat ? `rgb(${cat.color.join(",")})` : "#888";
}

const DocIcon = () => (
  <span className="doc-ic">
    <svg width="14" height="16" viewBox="0 0 14 16" fill="none">
      <path
        d="M1 2a1.5 1.5 0 0 1 1.5-1.5H9L13 4.5V14A1.5 1.5 0 0 1 11.5 15.5h-9A1.5 1.5 0 0 1 1 14V2Z"
        stroke="currentColor"
        strokeWidth="1.1"
      />
      <path d="M8.8.9V4.7h3.9" stroke="currentColor" strokeWidth="1.1" />
    </svg>
  </span>
);

function CatChips({ cats, config }: { cats: CategoryTotals; config: Config }) {
  const entries = Object.entries(cats);
  if (entries.length === 0) return <span className="no-tag">(no tag)</span>;
  return (
    <span className="cat-chips">
      {entries.map(([label, amt]) => (
        <span className="cat-chip" key={label}>
          <span className="dot" style={{ background: catColor(config, label) }} />
          <span>{label}</span>
          <span className="amt">{fmtMoney(amt)}</span>
        </span>
      ))}
    </span>
  );
}

function InvoiceCard({ result, config }: { result: CardResult; config: Config }) {
  const pages = result.pages ?? [];
  const multi = pages.length > 1;
  return (
    <article className="inv-card">
      <div className="inv-card-head">
        <span className={`status-dot ${result.status}`} />
        <span className="inv-name">{result.name}</span>
        {result.vendor && <span className="vendor-tag">{result.vendor}</span>}
      </div>
      {result.status === "ok" && (
        <div className="inv-card-body">
          {pages.map((cats, i) => (
            <div className="page-row" key={i}>
              <span className="page-label">Page {i + 1}</span>
              <CatChips cats={cats} config={config} />
            </div>
          ))}
          {multi && result.totals && (
            <div className="inv-total-row">
              <span className="page-label">Invoice</span>
              <CatChips cats={result.totals} config={config} />
            </div>
          )}
        </div>
      )}
      {result.status === "err" && (
        <div className="inv-card-body">
          <pre className="err-log">{(result.log ?? []).slice(1).join("\n")}</pre>
        </div>
      )}
      {result.status === "ok" && result.taggedName && result.url && (
        <div className="inv-card-foot">
          <span className="tagged-name">{result.taggedName}</span>
          <a className="dl-link" href={result.url} download={result.taggedName}>
            ↓ Download
          </a>
        </div>
      )}
    </article>
  );
}

export default function App() {
  const [config, setConfig] = useState<Config>(() => loadConfig());
  const [editorOpen, setEditorOpen] = useState(false);
  const [invoiceFiles, setInvoiceFiles] = useState<File[]>([]);
  const [masterFile, setMasterFile] = useState<File | null>(null);
  const [processing, setProcessing] = useState(false);
  const [results, setResults] = useState<CardResult[]>([]);
  const [masterDownload, setMasterDownload] = useState<Download | null>(null);
  const [masterError, setMasterError] = useState("");
  const [batchTime, setBatchTime] = useState("");
  const [dragOver, setDragOver] = useState(false);

  const invoiceInput = useRef<HTMLInputElement>(null);
  const masterInput = useRef<HTMLInputElement>(null);

  const addInvoices = useCallback((files: FileList | File[]) => {
    const pdfs = Array.from(files).filter((f) =>
      f.name.toLowerCase().endsWith(".pdf"),
    );
    setInvoiceFiles((prev) => {
      const names = new Set(prev.map((f) => f.name));
      return [...prev, ...pdfs.filter((f) => !names.has(f.name))];
    });
  }, []);

  const onDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragOver(false);
      addInvoices(e.dataTransfer.files);
    },
    [addInvoices],
  );

  async function run() {
    if (invoiceFiles.length === 0 || processing) return;
    setProcessing(true);
    setResults([]);
    setMasterDownload(null);
    setMasterError("");
    setBatchTime(new Date().toLocaleString());

    const finished: CardResult[] = [];
    for (const file of invoiceFiles) {
      setResults([...finished, { name: file.name, status: "busy" }]);
      const bytes = new Uint8Array(await file.arrayBuffer());
      const result = await processInvoice(file.name, bytes, config);
      finished.push({
        ...result,
        status: result.ok ? "ok" : "err",
        url:
          result.ok && result.taggedBytes
            ? toDownload(result.taggedName!, result.taggedBytes).url
            : undefined,
      });
      setResults([...finished]);
    }

    const tagged = finished.filter((r) => r.status === "ok" && r.taggedBytes);
    if (tagged.length > 0) {
      try {
        const masterBytes = masterFile
          ? new Uint8Array(await masterFile.arrayBuffer())
          : null;
        const updated = await appendToMaster(
          masterBytes,
          tagged.map((r) => r.taggedBytes!),
        );
        setMasterDownload(toDownload("master_invoices.pdf", updated));
      } catch (e) {
        setMasterError(
          `Could not update master PDF: ${e instanceof Error ? e.message : String(e)}`,
        );
      }
    }
    setProcessing(false);
  }

  function reset() {
    setInvoiceFiles([]);
    setMasterFile(null);
    setResults([]);
    setMasterDownload(null);
    setMasterError("");
    if (masterInput.current) masterInput.current.value = "";
  }

  return (
    <div className="app">
      <header className="masthead">
        <div>
          <p className="overline">Back of House</p>
          <h1 className="wordmark">Invoice Processor</h1>
          <p className="tagline">
            Tags PDF invoices by cost category — everything runs in your
            browser, no files are uploaded anywhere.
          </p>
        </div>
        <button className="btn ghost" onClick={() => setEditorOpen(true)}>
          Edit Categories
        </button>
      </header>

      <section
        className={`dropzone${dragOver ? " over" : ""}`}
        onDragOver={(e) => {
          e.preventDefault();
          setDragOver(true);
        }}
        onDragLeave={() => setDragOver(false)}
        onDrop={onDrop}
        onClick={() => invoiceInput.current?.click()}
      >
        <input
          ref={invoiceInput}
          type="file"
          accept=".pdf"
          multiple
          hidden
          onChange={(e) => {
            if (e.target.files) addInvoices(e.target.files);
            e.target.value = "";
          }}
        />
        {invoiceFiles.length === 0 ? (
          <div>
            <p className="drop-title">Drop invoice PDFs here</p>
            <p className="drop-hint">or click to choose files</p>
          </div>
        ) : (
          <div>
            <ul className="filelist">
              {invoiceFiles.map((f) => (
                <li key={f.name}>
                  <DocIcon />
                  <span className="fname">{f.name}</span>
                  <span className="fsize">{fmtSize(f.size)}</span>
                  <button
                    className="remove"
                    onClick={(e) => {
                      e.stopPropagation();
                      setInvoiceFiles((prev) => prev.filter((x) => x !== f));
                    }}
                  >
                    ✕
                  </button>
                </li>
              ))}
            </ul>
            <p className="add-more">Drop more files, or click to add</p>
          </div>
        )}
      </section>

      <section className="master-row">
        <span className="master-label">Master PDF to append to</span>
        <input
          ref={masterInput}
          type="file"
          accept=".pdf"
          hidden
          onChange={(e) => setMasterFile(e.target.files?.[0] ?? null)}
        />
        {masterFile ? (
          <span className="master-chip">
            {masterFile.name}
            <button
              className="remove"
              onClick={() => {
                setMasterFile(null);
                if (masterInput.current) masterInput.current.value = "";
              }}
            >
              ✕
            </button>
          </span>
        ) : (
          <button
            className="btn small ghost"
            onClick={() => masterInput.current?.click()}
          >
            Choose file — optional
          </button>
        )}
      </section>

      <section className="actions">
        <button
          className="btn primary"
          disabled={processing || invoiceFiles.length === 0}
          onClick={run}
        >
          {processing
            ? "Processing…"
            : `Process ${invoiceFiles.length || ""} Invoice${invoiceFiles.length === 1 ? "" : "s"}`}
        </button>
        <button className="btn ghost" disabled={processing} onClick={reset}>
          Clear
        </button>
      </section>

      {results.length > 0 && (
        <section>
          <div className="results-head">
            <h2 className="section-title">Results</h2>
            <span className="batch-meta">
              Batch: {results.length} file(s) — {batchTime}
            </span>
          </div>

          {masterDownload && (
            <div className="master-dl">
              <a
                className="btn outline-accent"
                href={masterDownload.url}
                download={masterDownload.name}
              >
                ↓ Updated {masterDownload.name}
              </a>
            </div>
          )}
          {masterError && <p className="error">{masterError}</p>}

          <div className="invoice-cards">
            {results.map((r) => (
              <InvoiceCard key={r.name} result={r} config={config} />
            ))}
          </div>
        </section>
      )}

      <p className="foot-note">
        Everything stays on this machine — nothing is uploaded.
      </p>

      {editorOpen && (
        <CategoryEditor
          config={config}
          onSave={(c) => {
            setConfig(c);
            setEditorOpen(false);
          }}
          onClose={() => setEditorOpen(false)}
        />
      )}
    </div>
  );
}
