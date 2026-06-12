import { useRef, useState } from "react";
import type { Category, Config } from "./lib/types";
import { defaultConfig, exportConfig, importConfig } from "./lib/config";

interface Props {
  config: Config;
  onSave: (config: Config) => void;
  onClose: () => void;
}

function rgbToHex([r, g, b]: [number, number, number]): string {
  return "#" + [r, g, b].map((v) => v.toString(16).padStart(2, "0")).join("");
}

function hexToRgb(hex: string): [number, number, number] {
  const h = hex.replace("#", "");
  return [
    parseInt(h.slice(0, 2), 16),
    parseInt(h.slice(2, 4), 16),
    parseInt(h.slice(4, 6), 16),
  ];
}

export default function CategoryEditor({ config, onSave, onClose }: Props) {
  const [cats, setCats] = useState<Category[]>(() =>
    structuredClone(config.categories),
  );
  const [selected, setSelected] = useState(0);
  const [newKeyword, setNewKeyword] = useState("");
  const [error, setError] = useState("");
  const importInput = useRef<HTMLInputElement>(null);

  const cat = cats[selected] as Category | undefined;

  function update(patch: Partial<Category>) {
    setCats((prev) =>
      prev.map((c, i) => (i === selected ? { ...c, ...patch } : c)),
    );
  }

  function addKeyword() {
    const kw = newKeyword.trim().toLowerCase();
    if (!kw || !cat) return;
    if (cat.keywords.includes(kw)) {
      setError(`"${kw}" is already in this category.`);
      return;
    }
    update({ keywords: [...cat.keywords, kw] });
    setNewKeyword("");
    setError("");
  }

  function save() {
    for (const c of cats) {
      if (!c.name.trim() || !c.code.trim()) {
        setError("Every category needs a name and a code.");
        return;
      }
    }
    onSave({ categories: cats });
  }

  function doExport() {
    const blob = new Blob([exportConfig({ categories: cats })], {
      type: "application/json",
    });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "categories.json";
    a.click();
    URL.revokeObjectURL(a.href);
  }

  async function doImport(file: File) {
    try {
      const next = importConfig(await file.text());
      setCats(next.categories);
      setSelected(0);
      setError("");
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  }

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <header className="modal-head">
          <h2>Edit Categories</h2>
          <div className="modal-tools">
            <button className="btn small ghost" onClick={doExport}>
              Export JSON
            </button>
            <button
              className="btn small ghost"
              onClick={() => importInput.current?.click()}
            >
              Import JSON
            </button>
            <input
              ref={importInput}
              type="file"
              accept=".json"
              hidden
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) doImport(f);
                e.target.value = "";
              }}
            />
            <button
              className="btn small ghost"
              onClick={() => {
                if (confirm("Reset all categories to the built-in defaults?")) {
                  setCats(defaultConfig().categories);
                  setSelected(0);
                }
              }}
            >
              Reset to Defaults
            </button>
          </div>
        </header>

        <div className="editor-body">
          <aside>
            <ul className="cat-list">
              {cats.map((c, i) => (
                <li
                  key={i}
                  className={i === selected ? "active" : ""}
                  onClick={() => setSelected(i)}
                >
                  <span
                    className="swatch"
                    style={{ background: rgbToHex(c.color) }}
                  />
                  <span className="cat-code">{c.code}</span>
                  <span>{c.name}</span>
                </li>
              ))}
            </ul>
            <div className="row">
              <button
                className="btn small outline-accent"
                onClick={() => {
                  setCats((prev) => [
                    ...prev,
                    {
                      code: "00000",
                      name: "New Category",
                      color: [180, 180, 180],
                      keywords: [],
                    },
                  ]);
                  setSelected(cats.length);
                }}
              >
                + Add
              </button>
              <button
                className="btn small danger"
                disabled={!cat}
                onClick={() => {
                  if (!cat) return;
                  if (!confirm(`Delete category "${cat.name}" (${cat.code})?`))
                    return;
                  setCats((prev) => prev.filter((_, i) => i !== selected));
                  setSelected(0);
                }}
              >
                − Delete
              </button>
            </div>
          </aside>

          {cat ? (
            <main>
              <div className="row meta">
                <label>
                  Name
                  <input
                    value={cat.name}
                    onChange={(e) => update({ name: e.target.value })}
                  />
                </label>
                <label>
                  Code
                  <input
                    value={cat.code}
                    onChange={(e) => update({ code: e.target.value })}
                  />
                </label>
                <label>
                  Color
                  <input
                    type="color"
                    value={rgbToHex(cat.color)}
                    onChange={(e) => update({ color: hexToRgb(e.target.value) })}
                  />
                </label>
              </div>

              <p className="hint">
                Keywords (case-insensitive, matched against line item
                descriptions — first matching category wins):
              </p>
              <ul className="kw-list">
                {cat.keywords.map((kw) => (
                  <li key={kw}>
                    <span>{kw}</span>
                    <button
                      className="remove"
                      onClick={() =>
                        update({
                          keywords: cat.keywords.filter((k) => k !== kw),
                        })
                      }
                    >
                      ✕
                    </button>
                  </li>
                ))}
              </ul>
              <div className="row">
                <input
                  className="field"
                  placeholder="new keyword…"
                  value={newKeyword}
                  onChange={(e) => setNewKeyword(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && addKeyword()}
                />
                <button className="btn small outline-accent" onClick={addKeyword}>
                  + Add Keyword
                </button>
              </div>
            </main>
          ) : (
            <main>
              <p className="hint">No category selected.</p>
            </main>
          )}
        </div>

        {error && <p className="error">{error}</p>}

        <footer className="modal-foot">
          <button className="btn ghost" onClick={onClose}>
            Cancel
          </button>
          <button className="btn primary" onClick={save}>
            Save Changes
          </button>
        </footer>
      </div>
    </div>
  );
}
