"use client";
import { useEffect, useRef, useState } from "react";
import type ExcelJS from "exceljs";
import Link from "next/link";
import WorkspaceTabs from "../workspace-tabs";
import { retainFormatting, exportAudit } from "@/lib/audit-export.mjs";
import { draft } from "@/lib/audit-draft";
import {
  isAuditSheet,
  readEntries,
  readDetails,
  prepareWorkbook,
  saveEntry,
  writeDetails,
  duplicateIds,
  totals,
  normalizeId,
} from "@/lib/audit.mjs";
import "./audit.css";

type Entry = ReturnType<typeof readEntries>[number];
const blank = { id: "", name: "", physical: "", system: "", cost: "" };
const money = (n: number) =>
  n.toLocaleString(undefined, { maximumFractionDigits: 2 });
export default function AuditEditor() {
  const workbook = useRef<ExcelJS.Workbook | null>(null),
    lock = useRef(false);
  const [busy, setBusy] = useState(true),
    [filename, setFilename] = useState(""),
    [sheetName, setSheetName] = useState(""),
    [sheets, setSheets] = useState<string[]>([]),
    [entries, setEntries] = useState<Entry[]>([]);
  const [details, setDetails] = useState({
      title: "",
      date: "",
      store: "",
      lastAudit: "",
      preparedBy: "",
    }),
    [form, setForm] = useState(blank),
    [editing, setEditing] = useState<string | null>(null),
    [search, setSearch] = useState("");
  const [notice, setNotice] = useState(""),
    [error, setError] = useState(""),
    [storageWarning, setStorageWarning] = useState(""),
    [detailsDirty, setDetailsDirty] = useState(false),
    [entryDirty, setEntryDirty] = useState(false);
  const dirty = detailsDirty || entryDirty;
  const fileInput = useRef<HTMLInputElement>(null);
  const sheet = () => workbook.current!.getWorksheet(sheetName)!;
  function display(w: ExcelJS.Workbook, name: string, file: string) {
    const s = w.getWorksheet(name)!;
    workbook.current = w;
    setSheets(w.worksheets.filter(isAuditSheet).map((s) => s.name));
    setSheetName(name);
    setFilename(file);
    setEntries(readEntries(s));
    setDetails(readDetails(s));
    setForm(blank);
    setEditing(null);
    setSearch("");
    setDetailsDirty(false);
    setEntryDirty(false);
  }
  async function parse(bytes: ArrayBuffer) {
    const { default: Excel } = await import("exceljs");
    const w = new Excel.Workbook();
    await w.xlsx.load(bytes);
    const supported = w.worksheets.filter(isAuditSheet);
    if (!supported.length)
      throw Error(
        "This workbook does not match your audit template. Use the XLSX with B/C, In H, In S, Balance and Unit Cost columns in both sections.",
      );
    await retainFormatting(w, bytes);
    prepareWorkbook(w);
    return w;
  }
  async function persist(
    w: ExcelJS.Workbook,
    name = sheetName,
    file = filename,
  ) {
    try {
      const buffer = await exportAudit(w);
      const bytes = new Uint8Array(buffer).slice().buffer;
      await draft("write", { bytes, filename: file, sheetName: name });
      setStorageWarning("");
    } catch {
      setStorageWarning(
        "Browser draft could not be saved. Download your workbook now to keep your changes.",
      );
    }
  }
  async function run(task: () => Promise<void>) {
    if (lock.current) return;
    lock.current = true;
    setBusy(true);
    setError("");
    try {
      await task();
    } catch (e) {
      setError(
        e instanceof Error ? e.message : "Unable to update this workbook.",
      );
    } finally {
      lock.current = false;
      setBusy(false);
    }
  }
  useEffect(() => {
    let active = true;
    void (async () => {
      try {
        const saved = await draft("read");
        if (saved) {
          const w = await parse(saved.bytes);
          if (active) {
            display(w, saved.sheetName, saved.filename);
            setNotice("Restored the saved draft from this browser.");
          }
        }
      } catch {
        if (active)
          setStorageWarning(
            "Browser draft is unavailable. Import your workbook and download changes before leaving.",
          );
      } finally {
        if (active) setBusy(false);
      }
    })();
    return () => {
      active = false;
    };
  }, []);
  useEffect(() => {
    const warn = (e: BeforeUnloadEvent) => {
      if (dirty) {
        e.preventDefault();
        e.returnValue = "";
      }
    };
    window.addEventListener("beforeunload", warn);
    return () => window.removeEventListener("beforeunload", warn);
  }, [dirty]);
  async function importFile(file: File) {
    if (!/\.xlsx$/i.test(file.name))
      throw Error(
        "Please choose an .xlsx workbook. Convert old .xls files in Excel first.",
      );
    if (file.size > 10 * 1024 * 1024)
      throw Error("Choose an audit workbook smaller than 10 MiB.");
    if (
      (workbook.current || dirty) &&
      !confirm(
        "Replace the current browser draft? Download it first if you need to keep it.",
      )
    )
      return;
    const w = await parse(await file.arrayBuffer());
    const name = w.worksheets.find(isAuditSheet)!.name;
    display(w, name, file.name);
    await persist(w, name, file.name);
    setNotice("Workbook imported. Search a Product ID before adding an entry.");
  }
  function edit(e: Entry) {
    if (
      entryDirty &&
      !confirm("Discard unsaved product changes and open this product?")
    )
      return;
    setEditing(e.key);
    setForm({
      id: e.id,
      name: e.name,
      physical: String(e.physical),
      system: String(e.system),
      cost: String(e.cost),
    });
    setEntryDirty(false);
    setNotice(
      "Editing the existing row. Quantities entered here replace its previous quantities.",
    );
    document
      .getElementById("product-entry")
      ?.scrollIntoView({ behavior: "smooth", block: "start" });
  }
  const duplicates = duplicateIds(entries),
    summary = totals(entries),
    query = search.trim().toLowerCase();
  const visible = entries.filter(
    (e) =>
      !query ||
      e.id.toLowerCase().includes(query) ||
      e.name.toLowerCase().includes(query),
  );
  const matches = entries.filter(
    (e) => normalizeId(e.id) === normalizeId(form.id) && e.key !== editing,
  );
  const day = details.date
    ? new Date(details.date + "T12:00:00").toLocaleDateString("en-GB", {
        weekday: "long",
      })
    : "";
  return (
    <div className="audit-page">
      <header
        className="audit-header"
        onClickCapture={(e) => {
          if (
            dirty &&
            (e.target as HTMLElement).closest("a") &&
            !confirm(
              "Leave without saving the current form changes? Saved workbook entries will stay in this browser.",
            )
          ) {
            e.preventDefault();
            e.stopPropagation();
          }
        }}
      >
        <Link href="/" className="audit-brand">
          <img src="/logo.svg" alt="" width="36" height="36" />
          FileBridge
        </Link>
        <WorkspaceTabs />
      </header>
      <main className="audit-main">
        <div className="audit-intro">
          <div>
            <div className="eyebrow">A SEPARATE WORKSPACE</div>
            <h1>Audit workbook</h1>
            <p>Search first. Enter once. Keep your Excel calculations.</p>
          </div>
          <button
            className="button"
            disabled={busy}
            onClick={() => fileInput.current?.click()}
          >
            Import workbook
          </button>
          <input
            type="file"
            accept=".xlsx"
            hidden
            ref={fileInput}
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) void run(() => importFile(f));
              e.target.value = "";
            }}
          />
        </div>
        <p className="audit-note">
          Your workbook stays in this browser, separate from File transfer.
          Saved edits are kept as a local draft. Download regularly to back up
          or move to another computer. Use one editor at a time; drafts do not
          sync between devices.
        </p>
        {error && (
          <div className="audit-alert" role="alert">
            {error}
          </div>
        )}
        {storageWarning && (
          <div className="audit-alert" role="alert">
            {storageWarning}
          </div>
        )}
        {notice && (
          <p className="audit-status" role="status">
            {notice}
          </p>
        )}
        {!workbook.current ? (
          <section className="audit-panel audit-welcome">
            <h2>
              {busy ? "Opening your draft…" : "Start with your audit workbook"}
            </h2>
            <p>
              Import your “Audit & Adjustments Report” XLSX. Its two product
              sections, formulas, and print layout will be retained. The source
              file is not overwritten.
            </p>
          </section>
        ) : (
          <>
            <section className="audit-panel">
              <h2>Report details</h2>
              <form
                onSubmit={(e) => {
                  e.preventDefault();
                  void run(async () => {
                    writeDetails(sheet(), details);
                    await persist(workbook.current!);
                    setDetailsDirty(false);
                    setNotice("Report details saved to browser draft.");
                  });
                }}
              >
                <div className="audit-fields">
                  <label>
                    Worksheet
                    <select
                      disabled={busy}
                      value={sheetName}
                      onChange={(e) => {
                        if (
                          dirty &&
                          !confirm(
                            "Discard unsaved form changes and switch worksheet?",
                          )
                        )
                          return;
                        const name = e.target.value;
                        display(workbook.current!, name, filename);
                        void run(() => persist(workbook.current!, name));
                      }}
                    >
                      {sheets.map((s) => (
                        <option key={s}>{s}</option>
                      ))}
                    </select>
                  </label>
                  {(
                    [
                      "title",
                      "date",
                      "store",
                      "lastAudit",
                      "preparedBy",
                    ] as const
                  ).map((key) => (
                    <label key={key}>
                      {
                        {
                          title: "Report title / products audited",
                          date: "Audit date",
                          store: "Store / department",
                          lastAudit: "Last audit label",
                          preparedBy: "Prepared / checked by",
                        }[key]
                      }
                      <input
                        disabled={busy}
                        required={key === "title" || key === "date"}
                        type={key === "date" ? "date" : "text"}
                        value={details[key]}
                        onChange={(e) => {
                          setDetails({ ...details, [key]: e.target.value });
                          setDetailsDirty(true);
                        }}
                      />
                    </label>
                  ))}
                  <label>
                    Download filename
                    <input
                      disabled={busy}
                      required
                      value={filename}
                      onChange={(e) => {
                        setFilename(e.target.value);
                        setDetailsDirty(true);
                      }}
                    />
                  </label>
                </div>
                <div className="audit-actions">
                  <span>{day}</span>
                  <button className="button primary" disabled={busy}>
                    Save report details
                  </button>
                  <button
                    type="button"
                    className="button"
                    disabled={busy}
                    onClick={() =>
                      void run(async () => {
                        if (entryDirty)
                          throw Error(
                            "Save or clear the product form before downloading.",
                          );
                        writeDetails(sheet(), details);
                        await persist(workbook.current!);
                        setDetailsDirty(false);
                        const buffer = await exportAudit(workbook.current!);
                        const url = URL.createObjectURL(
                          new Blob([new Uint8Array(buffer)], {
                            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                          }),
                        );
                        const a = document.createElement("a");
                        a.href = url;
                        a.download =
                          (filename
                            .trim()
                            .replace(/[\\/\x00-\x1f]/g, "_")
                            .replace(/\.xlsx$/i, "") || "Audit report") +
                          ".xlsx";
                        a.click();
                        setTimeout(() => URL.revokeObjectURL(url), 60000);
                        setNotice(
                          "Workbook downloaded. Open it in Excel to recalculate all formulas. Save product changes before downloading.",
                        );
                      })
                    }
                  >
                    Download Excel
                  </button>
                </div>
              </form>
            </section>
            <section className="audit-summary" aria-label="Audit totals">
              {[
                ["Products", entries.length],
                ["Shortage", money(summary.short)],
                ["Excess", money(summary.excess)],
                ["Net difference", money(summary.net)],
              ].map(([label, value]) => (
                <div key={label}>
                  <span>{label}</span>
                  <strong>{value}</strong>
                </div>
              ))}
            </section>
            {duplicates.length > 0 && (
              <div className="audit-alert" role="alert">
                This workbook already contains duplicate IDs:{" "}
                {duplicates.join(", ")}. Existing rows were retained. Edit the
                incorrect ID; new duplicates are blocked across both product
                sections.
              </div>
            )}
            <div className="audit-columns">
              <section className="audit-panel" id="product-entry">
                <h2>
                  {editing ? "Edit existing product" : "Add audited product"}
                </h2>
                <p>
                  In H = physically counted stock. In S = stock in your ERP.
                  Balance = In H − In S. Physical quantity and unit cost are
                  optional; leaving either blank saves 0.
                </p>
                <form
                  onSubmit={(e) => {
                    e.preventDefault();
                    void run(async () => {
                      const savedKey = saveEntry(sheet(), form, editing);
                      setEntries(readEntries(sheet()));
                      setSearch("");
                      await persist(workbook.current!);
                      setNotice(
                        editing
                          ? "Existing product updated. No new row added."
                          : `Product added to ${savedKey.startsWith("0:") ? "left" : "right"} section, Excel row ${savedKey.split(":")[1]}.`,
                      );
                      setForm(blank);
                      setEditing(null);
                      setEntryDirty(false);
                    });
                  }}
                >
                  <div className="audit-fields entry-fields">
                    {(
                      ["id", "name", "physical", "system", "cost"] as const
                    ).map((key) => (
                      <label key={key}>
                        {
                          {
                            id: "Product ID / barcode",
                            name: "Product name",
                            physical: "Physical quantity (In H)",
                            system: "ERP quantity (In S)",
                            cost: "Unit cost",
                          }[key]
                        }
                        <input
                          disabled={busy}
                          required={
                            key === "id" || key === "name" || key === "system"
                          }
                          placeholder={
                            key === "physical" || key === "cost"
                              ? "Optional — defaults to 0"
                              : undefined
                          }
                          inputMode={
                            key === "id"
                              ? "text"
                              : key === "name"
                                ? "text"
                                : "decimal"
                          }
                          type={
                            ["physical", "system", "cost"].includes(key)
                              ? "number"
                              : "text"
                          }
                          min={
                            ["physical", "system", "cost"].includes(key)
                              ? 0
                              : undefined
                          }
                          step="any"
                          value={form[key]}
                          onChange={(e) => {
                            setForm({ ...form, [key]: e.target.value });
                            setEntryDirty(true);
                          }}
                        />
                      </label>
                    ))}
                  </div>
                  {matches.length > 0 && (
                    <div className="audit-alert">
                      This ID already exists.{" "}
                      {matches.map((e) => (
                        <button
                          className="button"
                          type="button"
                          disabled={busy}
                          key={e.key}
                          onClick={() => edit(e)}
                        >
                          Open {e.name || e.id}
                        </button>
                      ))}
                    </div>
                  )}
                  <div className="audit-actions">
                    <button
                      className="button primary"
                      disabled={busy || matches.length > 0}
                    >
                      {busy
                        ? "Saving…"
                        : editing
                          ? "Update existing entry"
                          : "Add product"}
                    </button>
                    <button
                      type="button"
                      className="button"
                      disabled={busy}
                      onClick={() => {
                        if (
                          entryDirty &&
                          !confirm("Discard unsaved product changes?")
                        )
                          return;
                        setForm(blank);
                        setEditing(null);
                        setEntryDirty(false);
                      }}
                    >
                      Clear / new entry
                    </button>
                  </div>
                  <p className="audit-note">
                    {136 - entries.length} spaces left in this template.
                    Updating an entry replaces its quantities; add repeat counts
                    to the existing quantity yourself.
                  </p>
                </form>
              </section>
              <section className="audit-panel">
                <h2>Search existing products</h2>
                <label className="audit-search">
                  Product ID or name
                  <input
                    type="search"
                    placeholder="Search or scan a Product ID…"
                    value={search}
                    onChange={(e) => setSearch(e.target.value)}
                  />
                </label>
                <p>{visible.length} matching entries</p>
                <div className="audit-results">
                  {visible.map((e) => (
                    <article key={e.key}>
                      <div>
                        <strong>{e.name || "Unnamed product"}</strong>
                        <span>
                          ID: {e.id || "Missing ID"} ·{" "}
                          {e.offset === 0 ? "Left" : "Right"} section, row{" "}
                          {e.row}
                        </span>
                        <span>
                          Counted: {e.physical} · ERP: {e.system} · Cost:{" "}
                          {money(e.cost)}
                        </span>
                        <span>
                          Balance: {money(e.balance)} · Value: {money(e.value)}
                        </span>
                      </div>
                      <button
                        className="button"
                        disabled={busy}
                        onClick={() => edit(e)}
                      >
                        Edit
                      </button>
                    </article>
                  ))}
                  {!visible.length && (
                    <p>
                      No matching entry. Check the Product ID, then add it using
                      the form.
                    </p>
                  )}
                </div>
              </section>
            </div>
          </>
        )}
      </main>
    </div>
  );
}
