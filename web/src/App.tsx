import { useCallback, useRef, useState } from "react";
import type { Config } from "./lib/types";
import { loadConfig } from "./lib/config";
import { processInvoice, type ProcessResult } from "./lib/process";
import { appendToMaster } from "./lib/overlay";
import CategoryEditor from "./CategoryEditor";

interface Download {
  name: string;
  url: string;
}

function toDownload(name: string, bytes: Uint8Array): Download {
  const blob = new Blob([bytes as BlobPart], { type: "application/pdf" });
  return { name, url: URL.createObjectURL(blob) };
}

export default function App() {
  const [config, setConfig] = useState<Config>(() => loadConfig());
  const [editorOpen, setEditorOpen] = useState(false);
  const [invoiceFiles, setInvoiceFiles] = useState<File[]>([]);
  const [masterFile, setMasterFile] = useState<File | null>(null);
  const [processing, setProcessing] = useState(false);
  const [log, setLog] = useState<string[]>([]);
  const [downloads, setDownloads] = useState<Download[]>([]);
  const [masterDownload, setMasterDownload] = useState<Download | null>(null);
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
    if (invoiceFiles.length === 0) return;
    setProcessing(true);
    setDownloads([]);
    setMasterDownload(null);
    const lines: string[] = [
      `${"─".repeat(56)}`,
      `Batch: ${invoiceFiles.length} file(s)  —  ${new Date().toLocaleString()}`,
      `${"─".repeat(56)}`,
    ];
    setLog([...lines]);

    const results: ProcessResult[] = [];
    for (const file of invoiceFiles) {
      const bytes = new Uint8Array(await file.arrayBuffer());
      const result = await processInvoice(file.name, bytes, config);
      results.push(result);
      lines.push("", ...result.log);
      setLog([...lines]);
    }

    const tagged = results.filter((r) => r.ok && r.taggedBytes);
    setDownloads(tagged.map((r) => toDownload(r.taggedName!, r.taggedBytes!)));

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
        lines.push(
          "",
          masterFile
            ? `Master: appended ${tagged.length} invoice(s) to ${masterFile.name}`
            : `Master: created new master with ${tagged.length} invoice(s)`,
        );
      } catch (e) {
        lines.push("", `Master ERROR: ${e instanceof Error ? e.message : String(e)}`);
      }
    }

    const ok = results.filter((r) => r.ok).length;
    lines.push("", `${"─".repeat(56)}`, `Done — ${ok}/${results.length} succeeded.`);
    setLog([...lines]);
    setProcessing(false);
  }

  function reset() {
    setInvoiceFiles([]);
    setMasterFile(null);
    setDownloads([]);
    setMasterDownload(null);
    setLog([]);
  }

  return (
    <div className="app">
      <header>
        <div>
          <h1>Invoice Processor</h1>
          <p className="tagline">
            Tags PDF invoices by cost category — everything runs in your
            browser, no files are uploaded anywhere.
          </p>
        </div>
        <button className="btn purple" onClick={() => setEditorOpen(true)}>
          Edit Categories
        </button>
      </header>

      <section
        className={`dropzone ${dragOver ? "over" : ""}`}
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
          <p>
            <strong>Drop invoice PDFs here</strong> or click to choose files
          </p>
        ) : (
          <ul className="filelist">
            {invoiceFiles.map((f) => (
              <li key={f.name}>
                {f.name}
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
        )}
      </section>

      <section className="master-row">
        <label>
          Master PDF to append to (optional):
          <input
            ref={masterInput}
            type="file"
            accept=".pdf"
            onChange={(e) => setMasterFile(e.target.files?.[0] ?? null)}
          />
        </label>
        {masterFile && (
          <button
            className="remove"
            onClick={() => {
              setMasterFile(null);
              if (masterInput.current) masterInput.current.value = "";
            }}
          >
            ✕
          </button>
        )}
      </section>

      <section className="actions">
        <button
          className="btn green"
          disabled={processing || invoiceFiles.length === 0}
          onClick={run}
        >
          {processing ? "Processing…" : `Process ${invoiceFiles.length || ""} Invoice(s)`}
        </button>
        <button className="btn grey" disabled={processing} onClick={reset}>
          Clear
        </button>
      </section>

      {(downloads.length > 0 || masterDownload) && (
        <section className="downloads">
          <h2>Downloads</h2>
          {masterDownload && (
            <a
              className="btn purple"
              href={masterDownload.url}
              download={masterDownload.name}
            >
              ⬇ Updated {masterDownload.name}
            </a>
          )}
          <ul>
            {downloads.map((d) => (
              <li key={d.name}>
                <a href={d.url} download={d.name}>
                  ⬇ {d.name}
                </a>
              </li>
            ))}
          </ul>
        </section>
      )}

      {log.length > 0 && (
        <section className="logpanel">
          <pre>{log.join("\n")}</pre>
        </section>
      )}

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
