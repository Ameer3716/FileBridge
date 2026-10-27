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
