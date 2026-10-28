# FileBridge

A simple website for moving Excel files and other files between computers.
Open the same URL on both PCs. Upload on one, download on the other.
**There is no website login, signup, ChatGPT connection, or WhatsApp dependency.**

## Put it on GitHub and Vercel

1. Extract this ZIP. Create a GitHub repository and upload the contents of the `filebridge` folder. `package.json` must be in the repository root. Include the `app`, `lib`, `public`, and `tests` folders and configuration files.
2. In Vercel, choose **Add New → Project**, import the GitHub repository, and deploy. The framework is **Next.js**; use the default settings and Node.js 22 or later.
3. In that Vercel project's **Storage** tab, create and connect a **Blob** store. Choose **Private** storage. Connect the Production environment (and Preview if needed), keeping the default `BLOB_READ_WRITE_TOKEN` variable name.
4. **Redeploy** after connecting storage so the environment variable is available. In Project Settings → Deployment Protection, make sure **the production URL does not require Vercel Authentication**. Deployment protection is a hosting setting; the application itself has no login.
5. Open the **production URL** on both computers. Upload a file on the first PC. It appears on the second within 30 seconds; use Refresh to see it immediately.

The storage connection is a one-time owner setup. Visitors never sign into Vercel, Blob, GitHub, ChatGPT, or this website.

## Included features

- Drag and drop or browse; select multiple files.
- All files, recent uploads (last 7 days), and spreadsheet-only views.
- Search by filename; filter by uploaded date; sort newest, oldest, filename, or size.
- Server-recorded upload timestamps and original file-modified timestamps shown separately.
- Every displayed date and time uses **Pakistan Standard Time (PKT / UTC+5)**.
- Every upload has its own UUID. Identical names and identical content do not overwrite earlier uploads.
- Matching filenames show “Latest upload” and “Earlier upload”. “Latest” means latest uploaded, not necessarily latest edited; inspect the modified time when comparing revisions.
- Downloads preserve file contents and the original uploaded filename, including its extension. Duplicate uploads remain separate by their upload IDs.
- File details show exact byte size, upload ID, original modified time, and a SHA-256 fingerprint calculated by the browser before upload.
- Delete an individual upload from its details panel, with confirmation.
- Up to **25 MiB (26,214,400 bytes)** per file. Files upload directly to Blob, avoiding the serverless request-body upload limit. Large files use multipart uploads; downloads are streamed by the server.
- No automatic expiration. Files remain until deleted or the storage store is removed.

## Access model

This is deliberately an open, no-account file shelf. **Anyone with the website URL can list, upload, download, and delete files.** The site asks search engines not to index it, but that is not access control. Private Blob storage keeps the storage token server-side; the website exposes files without authentication by design. Do not put the token in browser code or commit it to GitHub.

Vercel storage and bandwidth are subject to your Vercel plan's quotas and pricing. A connected durable store is necessary: browser localStorage and Vercel's temporary function filesystem cannot transfer persistent files between PCs.

## Local development (optional)

```bash
npm ci
cp .env.example .env.local
# Put your Blob read/write token in .env.local, never in source files.
npm run dev
```

Open http://localhost:3000. Local upload completion callbacks are not required for the file list; Blob is the durable index.

```bash
npm test
npm run typecheck
npm run build
```

## Validation

The ZIP includes tests for duplicate-name isolation, Unicode filenames, filename sanitization, server timestamps, original download filenames and extensions, and cross-origin mutation rejection. Production build and TypeScript are checked before delivery. Live uploads/downloads require your connected Blob store and must be checked after deployment; no storage credentials are included.

## Project structure

- `app/file-shelf.tsx`: interface, upload progress, filtering, version labels, download actions.
- `app/api/upload/route.ts`: short-lived, size-limited direct-upload token.
- `app/api/files/route.ts`: shared file listing with full Blob pagination.
- `app/api/file/route.ts`: streamed downloads and deletion.
- `lib/files.mjs`: path metadata, unique download filenames, validation helpers.
- `tests/files.test.mjs`: data and timestamp tests.

## Official setup references

- https://vercel.com/docs/vercel-blob/client-upload
- https://vercel.com/docs/vercel-blob/private-storage
- https://vercel.com/docs/deployment-protection

## Audit workbook tab

Open **Audit workbook** (`/audit`) separately from File transfer. Import the two-panel XLSX audit template (headers on row 4; entries in rows 5–72). The original file is never overwritten or automatically uploaded. No sample audit workbook is bundled with the site.

- Edit report title/product group, audit date (weekday shown), store/department, last-audit label, preparer, and download filename.
- Search by Product ID/barcode or name. New IDs are checked across both panels, ignoring surrounding spaces and letter case. Existing duplicates are reported, not silently removed. Leading-zero text IDs remain text.
- Enter product name and ERP/system quantity (In S). Physical quantity (In H) and unit cost are optional and default to zero when blank. Edit an existing row to replace its quantities; the app does not automatically add repeated counts.
- Original Excel formulas remain in place. The browser shows balance/value and shortage/excess/net totals for this template; Excel recalculates on opening the export. Import rejects other sheet layouts. The 136 entry slots are fixed so adding products never overwrites the report totals.
- Save the product form and report details, then download XLSX. Saved workbook changes are retained in IndexedDB on this browser. Storage failures are reported; download regularly for backup. A new import replaces the local draft after confirmation. Unsaved form fields are not part of the export.
- Drafts are local, not a shared database. Use one editor at a time. To move devices, download the edited workbook and transfer it using the existing File transfer tab. Duplicate checks apply to the currently imported worksheet, not other workbooks or separate device drafts.

## Product duplicate checks in File transfer

Uploaded `.xlsx` files (up to 10 MiB for checking) are automatically scanned in the **Product duplicate check** panel at the bottom. For an existing shared file, use **Check products**; **Download** also checks it before handing out a copy. Other file types continue to transfer normally; convert old `.xls` workbooks to `.xlsx` in Excel to check them.

The checker supports the supplied two-panel audit template and product tables with recognizable ID/barcode/B/C and Product Name headings in the first 30 rows. IDs and names are checked independently across all recognized product tables, including across worksheets. It reports sheet names and Excel row numbers. Blank values are ignored, leading-zero IDs are significant, and name comparison ignores case, Unicode width, and extra whitespace. Unrecognized sheets are explicitly marked as unchecked; malformed files never receive a clean result.

Edit an incorrect ID/name or explicitly clear an extra product entry (including its quantities/cost, after confirmation). Entries are never automatically merged and quantities are not added together. All duplicate groups must be resolved before **Download checked copy** is enabled. Formula-based or merged identities must be corrected in Excel. Open corrected copies in Excel to recalculate formulas. Corrections exist in the current browser session and downloaded copy; original shared uploads remain unchanged. Upload the corrected copy to share it with another computer.
