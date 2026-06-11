// Invoice Processor — redesigned App (mirrors web/src/App.tsx structure)
const { useCallback, useEffect, useRef, useState } = React;

window.DEFAULT_CONFIG = {
  categories: [
    { code: "52000", name: "Merch",   color: [0, 191, 255],   keywords: ["yerba mate", "red bull", "sticker"] },
    { code: "50900", name: "F&B",     color: [255, 200, 0],   keywords: ["croissant", "kombucha", "gelato"] },
    { code: "53100", name: "Kitchen", color: [153, 51, 255],  keywords: ["cup lid", "straw", "napkin"] },
    { code: "61600", name: "cafe",    color: [100, 200, 100], keywords: ["sanitizer", "detergent", "rinse aid"] },
  ],
};

// ── Demo data (mirrors a real batch run) ─────────────────────────────
const SAMPLE_FILES = [
  { name: "Invoice.pdf", size: 182_044 },
  { name: "Invoice2.pdf", size: 421_380 },
  { name: "Invoice3.pdf", size: 396_217 },
];

const SAMPLE_RESULTS = {
  "Invoice.pdf": {
    vendor: "sysco",
    pages: [{ "50900 F&B": 169.43 }],
    totals: { "50900 F&B": 169.43 },
  },
  "Invoice2.pdf": {
    vendor: "sysco",
    pages: [
      { "50900 F&B": 1623.07, "52000 Merch": 316.94 },
      { "50900 F&B": 6.5 },
    ],
    totals: { "50900 F&B": 1629.57, "52000 Merch": 316.94 },
  },
  "Invoice3.pdf": {
    vendor: "sysco",
    pages: [
      { "50900 F&B": 1906.57, "52000 Merch": 85.97 },
      { "52000 Merch": 25.95, "50900 F&B": 65.27 },
    ],
    totals: { "50900 F&B": 1971.84, "52000 Merch": 111.92 },
  },
};

const TWEAK_DEFAULTS = /*EDITMODE-BEGIN*/{
  "accent": "#C4885A",
  "headingFont": "Garamond",
  "resultsView": "Cards"
}/*EDITMODE-END*/;

// ── Helpers ──────────────────────────────────────────────────────────
function fmtMoney(v) {
  return "$" + v.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function fmtSize(bytes) {
  return bytes > 1024 * 1024
    ? (bytes / (1024 * 1024)).toFixed(1) + " MB"
    : Math.round(bytes / 1024) + " KB";
}
function timestamp() {
  const d = new Date();
  const p = (n, w = 2) => String(n).padStart(w, "0");
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
}
function catColor(config, label) {
  const code = label.split(" ")[0];
  const cat = config.categories.find((c) => c.code === code);
  return cat ? "rgb(" + cat.color.join(",") + ")" : "#888";
}
function demoBlobUrl(name) {
  return URL.createObjectURL(new Blob(["Demo placeholder for " + name], { type: "text/plain" }));
}
function simulateResult(file) {
  const known = SAMPLE_RESULTS[file.name];
  if (known) return known;
  // Deterministic plausible result for arbitrary dropped files
  const amt = Math.round((file.size % 90000) + 4200) / 100;
  return { vendor: "sysco", pages: [{ "50900 F&B": amt }], totals: { "50900 F&B": amt } };
}

const DocIcon = () => (
  <span className="doc-ic">
    <svg width="14" height="16" viewBox="0 0 14 16" fill="none">
      <path d="M1 2a1.5 1.5 0 0 1 1.5-1.5H9L13 4.5V14A1.5 1.5 0 0 1 11.5 15.5h-9A1.5 1.5 0 0 1 1 14V2Z" stroke="currentColor" strokeWidth="1.1"></path>
      <path d="M8.8.9V4.7h3.9" stroke="currentColor" strokeWidth="1.1"></path>
    </svg>
  </span>
);

// ── Result card ──────────────────────────────────────────────────────
function CatChips({ cats, config }) {
  const entries = Object.entries(cats);
  if (entries.length === 0) return <span className="no-tag">(no tag)</span>;
  return (
    <span className="cat-chips">
      {entries.map(([label, amt]) => (
        <span className="cat-chip" key={label}>
          <span className="dot" style={{ background: catColor(config, label) }}></span>
          <span>{label}</span>
          <span className="amt">{fmtMoney(amt)}</span>
        </span>
      ))}
    </span>
  );
}

function InvoiceCard({ result, config }) {
  const multi = result.pages && result.pages.length > 1;
  return (
    <article className="inv-card">
      <div className="inv-card-head">
        <span className={"status-dot " + result.status}></span>
        <span className="inv-name">{result.name}</span>
        {result.vendor && <span className="vendor-tag">{result.vendor}</span>}
      </div>
      {result.status !== "busy" && (
        <div className="inv-card-body">
          {result.pages.map((cats, i) => (
            <div className="page-row" key={i}>
              <span className="page-label">Page {i + 1}</span>
              <CatChips cats={cats} config={config}></CatChips>
            </div>
          ))}
          {multi && (
            <div className="inv-total-row">
              <span className="page-label">Invoice</span>
              <CatChips cats={result.totals} config={config}></CatChips>
            </div>
          )}
        </div>
      )}
      {result.taggedName && (
        <div className="inv-card-foot">
          <span className="tagged-name">{result.taggedName}</span>
          <a className="dl-link" href={result.url} download={result.taggedName}>↓ Download</a>
        </div>
      )}
    </article>
  );
}

// ── Terminal log builder (original format, for the Terminal tweak) ──
function buildLog(results, batchTime) {
  const bar = "─".repeat(56);
  const lines = [bar, `Batch: ${results.length} file(s)  —  ${batchTime}`, bar];
  for (const r of results) {
    lines.push("", `> ${r.name}`, `  Vendor: ${r.vendor}`);
    r.pages.forEach((cats, i) => {
      const s = Object.entries(cats).map(([k, v]) => `${k}  ${fmtMoney(v)}`).join("  |  ");
      lines.push(`  Page ${i + 1}: ${s || "(no tag)"}`);
    });
    if (r.pages.length > 1) {
      const s = Object.entries(r.totals).map(([k, v]) => `${k}  ${fmtMoney(v)}`).join("  |  ");
      lines.push(`  Invoice (p1–${r.pages.length}): ${s}`);
    }
    lines.push(`  Tagged: ${r.taggedName}`);
  }
  lines.push("", `Master: created new master with ${results.length} invoice(s)`);
  lines.push("", bar, `Done — ${results.length}/${results.length} succeeded.`);
  return lines.join("\n");
}

// ── App ──────────────────────────────────────────────────────────────
function App() {
  const [t, setTweak] = useTweaks(TWEAK_DEFAULTS);
  const [config, setConfig] = useState(window.DEFAULT_CONFIG);
  const [editorOpen, setEditorOpen] = useState(false);
  const [invoiceFiles, setInvoiceFiles] = useState(SAMPLE_FILES);
  const [masterFile, setMasterFile] = useState(null);
  const [processing, setProcessing] = useState(false);
  const [results, setResults] = useState([]);
  const [masterDownload, setMasterDownload] = useState(null);
  const [batchTime, setBatchTime] = useState("");
  const [dragOver, setDragOver] = useState(false);

  const invoiceInput = useRef(null);
  const masterInput = useRef(null);

  useEffect(() => {
    const root = document.documentElement;
    root.style.setProperty("--accent", t.accent);
    root.style.setProperty(
      "--serif",
      t.headingFont === "Marcellus" ? '"Marcellus", Georgia, serif' : '"Cormorant Garamond", Georgia, serif'
    );
  }, [t.accent, t.headingFont]);

  const addInvoices = useCallback((files) => {
    const pdfs = Array.from(files).filter((f) => f.name.toLowerCase().endsWith(".pdf"));
    setInvoiceFiles((prev) => {
      const names = new Set(prev.map((f) => f.name));
      return [...prev, ...pdfs.filter((f) => !names.has(f.name))];
    });
  }, []);

  const onDrop = useCallback((e) => {
    e.preventDefault();
    setDragOver(false);
    addInvoices(e.dataTransfer.files);
  }, [addInvoices]);

  async function run() {
    if (invoiceFiles.length === 0 || processing) return;
    setProcessing(true);
    setResults([]);
    setMasterDownload(null);
    setBatchTime(new Date().toLocaleString());
    const ts = timestamp();
    const finished = [];

    for (const file of invoiceFiles) {
      // show busy card
      setResults([...finished, { name: file.name, status: "busy", pages: [] }]);
      await new Promise((r) => setTimeout(r, 650));
      const sim = simulateResult(file);
      const taggedName = file.name.replace(/\.pdf$/i, "") + "_tagged_" + ts + ".pdf";
      finished.push({
        name: file.name,
        status: "ok",
        vendor: sim.vendor,
        pages: sim.pages,
        totals: sim.totals,
        taggedName,
        url: demoBlobUrl(taggedName),
      });
      setResults([...finished]);
    }

    setMasterDownload({ name: "master_invoices.pdf", url: demoBlobUrl("master_invoices.pdf") });
    setProcessing(false);
  }

  function reset() {
    setInvoiceFiles([]);
    setMasterFile(null);
    setResults([]);
    setMasterDownload(null);
  }

  const done = results.filter((r) => r.status === "ok");

  return (
    <div className="app" data-screen-label="Invoice Processor">
      <header className="masthead">
        <div>
          <p className="overline">Back of House</p>
          <h1 className="wordmark">Invoice Processor</h1>
          <p className="tagline">
            Tags PDF invoices by cost category — everything runs in your browser,
            no files are uploaded anywhere.
          </p>
        </div>
        <button className="btn ghost" onClick={() => setEditorOpen(true)}>Edit Categories</button>
      </header>

      <section
        className={"dropzone" + (dragOver ? " over" : "")}
        onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
        onDragLeave={() => setDragOver(false)}
        onDrop={onDrop}
        onClick={() => invoiceInput.current && invoiceInput.current.click()}
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
        ></input>
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
                  <DocIcon></DocIcon>
                  <span className="fname">{f.name}</span>
                  <span className="fsize">{fmtSize(f.size)}</span>
                  <button
                    className="remove"
                    onClick={(e) => {
                      e.stopPropagation();
                      setInvoiceFiles((prev) => prev.filter((x) => x !== f));
                    }}
                  >✕</button>
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
          onChange={(e) => setMasterFile((e.target.files && e.target.files[0]) || null)}
        ></input>
        {masterFile ? (
          <span className="master-chip">
            {masterFile.name}
            <button
              className="remove"
              onClick={() => {
                setMasterFile(null);
                if (masterInput.current) masterInput.current.value = "";
              }}
            >✕</button>
          </span>
        ) : (
          <button className="btn small ghost" onClick={() => masterInput.current && masterInput.current.click()}>
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
          {processing ? "Processing…" : "Process " + (invoiceFiles.length || "") + " Invoice" + (invoiceFiles.length === 1 ? "" : "s")}
        </button>
        <button className="btn ghost" disabled={processing} onClick={reset}>Clear</button>
        <span className="demo-note">demo — processing is simulated in this mock</span>
      </section>

      {results.length > 0 && (
        <section>
          <div className="results-head">
            <h2 className="section-title">Results</h2>
            <span className="batch-meta">Batch: {results.length} file(s) — {batchTime}</span>
          </div>

          {masterDownload && (
            <div className="master-dl">
              <a className="btn outline-accent" href={masterDownload.url} download={masterDownload.name}>
                ↓ Updated {masterDownload.name}
              </a>
            </div>
          )}

          {t.resultsView === "Terminal" ? (
            done.length > 0 && (
              <div className="logpanel">
                <pre>{buildLog(done, batchTime)}</pre>
              </div>
            )
          ) : (
            <div className="invoice-cards">
              {results.map((r) => (
                <InvoiceCard key={r.name} result={r} config={config}></InvoiceCard>
              ))}
            </div>
          )}
        </section>
      )}

      <p className="foot-note">Everything stays on this machine — nothing is uploaded.</p>

      {editorOpen && (
        <CategoryEditor
          config={config}
          onSave={(c) => { setConfig(c); setEditorOpen(false); }}
          onClose={() => setEditorOpen(false)}
        ></CategoryEditor>
      )}

      <TweaksPanel>
        <TweakSection label="Theme"></TweakSection>
        <TweakColor
          label="Accent"
          value={t.accent}
          options={["#C4885A", "#8FA382", "#A8946B", "#B85C43"]}
          onChange={(v) => setTweak("accent", v)}
        ></TweakColor>
        <TweakRadio
          label="Heading face"
          value={t.headingFont}
          options={["Garamond", "Marcellus"]}
          onChange={(v) => setTweak("headingFont", v)}
        ></TweakRadio>
        <TweakSection label="Results"></TweakSection>
        <TweakRadio
          label="Log style"
          value={t.resultsView}
          options={["Cards", "Terminal"]}
          onChange={(v) => setTweak("resultsView", v)}
        ></TweakRadio>
      </TweaksPanel>
    </div>
  );
}

ReactDOM.createRoot(document.getElementById("root")).render(<App></App>);
