"use client";
import { useEffect, useState } from "react";
import {
  scanProducts,
  fixProduct,
  clearProduct,
  exportCheckedProducts,
} from "@/lib/product-check.mjs";
import type { loadProductCheck } from "@/lib/product-check.mjs";
export type ProductReview = {
  key: string;
  name: string;
  bytes?: ArrayBuffer;
  session?: Awaited<ReturnType<typeof loadProductCheck>>;
  status: "checking" | "ready" | "error";
  error?: string;
};
function save(bytes: ArrayBuffer | Uint8Array, name: string) {
  const url = URL.createObjectURL(
    new Blob([new Uint8Array(bytes)], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
  );
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}
export default function ProductCheckPanel({
  reviews,
  onChange,
  activeKey,
  onSelect,
}: {
  reviews: ProductReview[];
  activeKey: string;
  onSelect: (key: string) => void;
  onChange: (r: ProductReview) => void;
}) {
  const [editing, setEditing] = useState(""),
    [id, setId] = useState(""),
    [name, setName] = useState(""),
    [error, setError] = useState(""),
    [busy, setBusy] = useState(false);
  const review = reviews.find((r) => r.key === activeKey) ?? reviews.at(-1),
    scan = review?.session ? scanProducts(review.session) : null;
  useEffect(() => {
    setEditing("");
    setError("");
  }, [review?.key]);
  const affected =
    scan?.rows.filter((row) =>
      scan.groups.some((g) => g.rows.some((r) => r.key === row.key)),
    ) ?? [];
  return (
    <section
      className="product-check"
      id="product-check"
      aria-labelledby="product-check-title"
    >
      <h2 id="product-check-title">Product duplicate check</h2>
      <p>
        XLSX uploads are checked for repeated Product IDs and product names.
        Review duplicates here before downloading. Other file types transfer
        normally.
      </p>
      {!reviews.length ? (
        <p>
          Upload an XLSX file, or choose “Check products” beside an existing
          Excel file.
        </p>
      ) : (
        <>
          <label>
            File to review
            <select
              value={review?.key}
              disabled={busy}
              onChange={(e) => {
                onSelect(e.target.value);
                setEditing("");
                setError("");
              }}
            >
              {reviews.map((r) => (
                <option key={r.key} value={r.key}>
                  {r.name} —{" "}
                  {r.status === "checking"
                    ? "Checking…"
                    : r.status === "error"
                      ? "Not checked"
                      : r.session && scanProducts(r.session).groups.length
                        ? "Duplicates found"
                        : "No duplicates in checked tables"}
                </option>
              ))}
            </select>
          </label>
          {review?.status === "checking" && (
            <p role="status">Checking workbook…</p>
          )}
          {review?.status === "error" && (
            <div className="product-warning" role="alert">
              Not checked: {review.error}
              {review.bytes && (
                <button
                  className="button"
                  onClick={() => save(review.bytes!, review.name)}
                >
                  Download original (unchecked)
                </button>
              )}
            </div>
          )}
          {scan && (
            <>
              <p role="status">
                Checked {scan.rows.length} product entries in{" "}
                {scan.supported.join(", ")}.{" "}
                {scan.groups.length
                  ? `${scan.groups.length} duplicate groups found.`
                  : "No duplicate IDs or names in the checked tables."}
              </p>
              {scan.skipped.length > 0 && (
                <p className="product-warning">
                  Partial check: these sheets have no recognized product table
                  and were not checked: {scan.skipped.join(", ")}. Add Product
                  ID and Product Name headings in Excel if they contain product
                  rows.
                </p>
              )}
              {scan.groups.map((g) => (
                <div className="product-warning" key={g.field + g.value}>
                  <strong>
                    Duplicate {g.field === "id" ? "Product ID" : "product name"}
                    : {g.rows[0][g.field as "id" | "name"]}
                  </strong>
                  <p>
                    {g.rows
                      .map(
                        (r) =>
                          `${r.sheetName}, row ${r.row} (${r.idCol < 9 ? "left" : "right"} columns)`,
                      )
                      .join(" · ")}
                  </p>
                </div>
              ))}
              <div className="product-rows">
                {affected.map((row) => (
                  <article key={row.key}>
                    <div>
                      <strong>{row.name || "Unnamed product"}</strong>
                      <p>
                        ID: {row.id || "Missing ID"} · {row.sheetName}, Excel
                        row {row.row}
                      </p>
                    </div>
                    {editing === row.key ? (
                      <form
                        onSubmit={(e) => {
                          e.preventDefault();
                          try {
                            fixProduct(review!.session!, row.key, id, name);
                            onChange({ ...review! });
                            setEditing("");
                            setError("");
                          } catch (e) {
                            setError(
                              e instanceof Error
                                ? e.message
                                : "Could not update product.",
                            );
                          }
                        }}
                      >
                        <label>
                          Correct Product ID
                          <input
                            required
                            value={id}
                            onChange={(e) => setId(e.target.value)}
                          />
                        </label>
                        <label>
                          Correct product name
                          <input
                            required
                            value={name}
                            onChange={(e) => setName(e.target.value)}
                          />
                        </label>
                        <button className="button primary">
                          Save correction
                        </button>
                        <button
                          className="button"
                          type="button"
                          onClick={() => setEditing("")}
                        >
                          Cancel
                        </button>
                      </form>
                    ) : (
                      <div className="product-buttons">
                        <button
                          className="button"
                          disabled={busy}
                          onClick={() => {
                            setEditing(row.key);
                            setId(row.id);
                            setName(row.name);
                            setError("");
                          }}
                        >
                          Edit row {row.row}
                        </button>
                        <button
                          className="button danger"
                          disabled={busy}
                          onClick={() => {
                            if (
                              !confirm(
                                `Clear ${row.name || row.id} from ${row.sheetName}, row ${row.row}? Its quantities and cost will be cleared. Formulas remain. Counts are not merged into another entry.`,
                              )
                            )
                              return;
                            try {
                              clearProduct(review!.session!, row.key);
                              onChange({ ...review! });
                              setError("");
                            } catch (e) {
                              setError(
                                e instanceof Error
                                  ? e.message
                                  : "Could not clear entry.",
                              );
                            }
                          }}
                        >
                          Clear entry
                        </button>
                      </div>
                    )}
                  </article>
                ))}
              </div>
              <p>
                Names are compared without letter case, extra spaces, or
                Unicode-width differences. IDs ignore letter case and
                surrounding spaces; leading zeros remain significant. Check
                which product is correct—entries and counts are never merged
                automatically.
              </p>
              <button
                className="button primary"
                disabled={busy || scan.groups.length > 0 || !!editing}
                onClick={async () => {
                  setBusy(true);
                  setError("");
                  try {
                    save(
                      await exportCheckedProducts(review!.session!),
                      review!.name,
                    );
                  } catch (e) {
                    setError(
                      e instanceof Error
                        ? e.message
                        : "Could not export workbook.",
                    );
                  } finally {
                    setBusy(false);
                  }
                }}
              >
                {busy
                  ? "Preparing…"
                  : scan.skipped.length
                    ? "Download copy (partial check)"
                    : "Download checked copy"}
              </button>
              <p>
                Corrections apply to this browser session and download. The
                original upload stays unchanged. Upload the corrected copy to
                share it. Open it in Excel to recalculate formulas.
              </p>
            </>
          )}
        </>
      )}
      {error && (
        <p className="product-warning" role="alert">
          {error}
        </p>
      )}
    </section>
  );
}
