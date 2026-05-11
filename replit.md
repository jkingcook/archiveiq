# LetterVault Pro

A full-stack archival document processing app for Lyndon Watson Cook (LWC). Uses Claude AI to analyze archival images, extract metadata, transcribe text, and generate professional `.docx` exports.

## Run & Operate

- `pnpm --filter @workspace/lettervault run dev` — run the LetterVault server (port 20897) — **this is the main app**
- `pnpm --filter @workspace/api-server run dev` — run the legacy API server (port 8080, path `/api-backend`)
- `pnpm run typecheck` — full typecheck across all packages
- `pnpm run build` — typecheck + build all packages

## Stack

- pnpm workspaces, Node.js 24, TypeScript 5.9
- Express 5 server with vanilla HTML/CSS/JS frontend (`src/public/index.html`)
- Claude AI: `claude-sonnet-4-5` via `@anthropic-ai/sdk`
- `.docx` generation: `docx` library
- ZIP exports: `archiver` v8 (`ZipArchive` API)
- Image uploads: `multer` (memory storage, max 20MB)
- Logging: `pino` + `pino-http`
- Build: `esbuild` (ESM bundle)

## Where things live

- `artifacts/lettervault/` — the publishable LetterVault Pro web artifact
  - `src/index.ts` — server entry point
  - `src/app.ts` — Express app with static serving + `/api` router
  - `src/lib/logger.ts` — pino logger singleton
  - `src/routes/lettervault.ts` — all LetterVault API routes
  - `src/routes/health.ts` — `/api/healthz` health check
  - `src/public/index.html` — full frontend (4 tabs, vanilla JS)
  - `build.mjs` — esbuild config that bundles server + copies `public/`
- `artifacts/api-server/` — legacy API server (kind=api, not publishable), moved to `/api-backend`
- Generated `.docx` files stored in `/tmp/output/items/` and `/tmp/output/analysis/`

## Architecture decisions

- Single Express artifact serves both the API (`/api/*`) and the frontend static files — avoids the `kind=api` publishing restriction
- In-memory store (`itemStore[]`, `sessionStore`) for items — no database needed for archival session work
- Items processed one-at-a-time with 1s delay between Claude API calls to avoid rate limits
- Two-attempt retry loop on Claude responses in case JSON parsing fails
- esbuild bundles everything into a single `dist/index.mjs`, with `src/public/` copied to `dist/public/`

## Product

- **Tab 1 — Setup**: Paste and save a processing protocol (up to 10,000 chars), edit the JSON schema for items (Layer 1) and group analyses (Layer 2)
- **Tab 2 — Process Items**: Drag/drop archival images (JPG, PNG, WEBP), process with Claude AI one by one, download individual `.docx` reports
- **Tab 3 — Analyze Collection**: Group processed items (by year, decade, sender, type, era, etc.), run Claude synthesis, download group reports + master report
- **Tab 4 — All Exports**: Download any generated `.docx` or ZIP of all items/analysis

## User preferences

- App name: LetterVault Pro
- Client: Lyndon Watson Cook (LWC)
- AI model: `claude-sonnet-4-5` (not claude-sonnet-4-20250514 — that name caused API errors)
- Design: warm cream/gold/brown archival aesthetic, Georgia serif headings

## Gotchas

- `ANTHROPIC_API_KEY` must be set in Replit Secrets — the app shows a modal and status badge if missing
- `archiver` v8 uses `new ZipArchive()` not `archiver("zip")` — do not downgrade
- The old `api-server` artifact has `kind=api` and cannot be published — lettervault (kind=web) is the publishable one
- Do not delete `src/lib/logger.ts` when cleaning scaffold files — it's created manually and not from scaffold

## Pointers

- See the `pnpm-workspace` skill for workspace structure, TypeScript setup, and package details
