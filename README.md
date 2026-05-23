# DocuCaption Pro

DocuCaption Pro is a free Word document automation tool. It helps users upload a `.docx`, detect figures and tables, add missing Word-native captions, avoid duplicate captions, insert/update Word fields, and export a corrected Word copy.

## Core Features

- Upload `.docx` files.
- Detect figures from `w:drawing` and `w:pict`.
- Detect Word tables from `w:tbl`.
- Add missing captions using Word-native `SEQ Figure` and `SEQ Table` fields.
- Avoid duplicate captions near existing figure/table captions.
- Insert or normalize:
  - Table of Contents: `TOC \o "1-3" \h \z \u`
  - Table of Figures: `TOC \h \z \c "Figure"`
  - Table of Tables: `TOC \h \z \c "Table"`
- Show a document audit summary.
- Generate editable draft captions through the backend.
- Export as `originalName_captioned.docx` without overwriting the original.
- Prompt Word to update fields through `word/settings.xml`.

## Access Model

The app is free.

The backend still uses generous monthly usage counters to protect the caption generation service from accidental overuse:

- 1000 documents per month
- 1000 exports per month
- 5000 generated captions per month
- 200 figures/tables per document

These values are configured in [src/lib/access.ts](</C:/Users/User/Downloads/doccaptioner/src/lib/access.ts>).

## Supabase Setup

Create a Supabase project, enable Email and Google providers, then add `http://localhost:3000` and your production domain to Auth redirect URLs.

Run [supabase/schema.sql](</C:/Users/User/Downloads/doccaptioner/supabase/schema.sql>) in the Supabase SQL editor. It creates:

- `profiles`
- `documents`
- `usage_logs`
- private Storage bucket `docucaption-documents`
- indexes for account and usage lookups
- an Auth trigger that creates a `profiles` row whenever a new Supabase Auth user is created
- owner-only Storage policies for files stored under `{user_id}/...`
- service-role-only functions for atomic document/export quota and caption credit updates

The RLS policies allow signed-in users to read their own profile, documents, and usage logs. Client-side writes are intentionally not allowed for usage or caption credit fields. The backend writes those fields with `SUPABASE_SERVICE_ROLE_KEY`.

Profile creation is handled in two places:

- Supabase trigger `public.handle_new_auth_user()` creates `public.profiles` immediately after a new `auth.users` row is created.
- The backend `GET /api/account` route also ensures a profile exists after verifying a Supabase session.

Signed-in uploads are placed in the private `docucaption-documents` Storage bucket under `{user_id}/uploads/...` while the document is being processed. By default the frontend deletes that object immediately after the processing record is created.

Set `VITE_PERSIST_UPLOADED_DOCS="true"` only if you intentionally want to keep uploaded files in Supabase Storage for a production workflow.

Required environment variables:

```bash
VITE_SUPABASE_URL="https://your-project.supabase.co"
VITE_SUPABASE_ANON_KEY="your_public_anon_key"
SUPABASE_URL="https://your-project.supabase.co"
SUPABASE_SERVICE_ROLE_KEY="your_service_role_key"
VITE_PERSIST_UPLOADED_DOCS="false"
```

Frontend-safe variables:

- `VITE_SUPABASE_URL`
- `VITE_SUPABASE_ANON_KEY`

Backend-only variables:

- `SUPABASE_URL`
- `SUPABASE_SERVICE_ROLE_KEY`
- `GEMINI_API_KEY`

## Gemini Backend Setup

The Gemini key is never exposed in the frontend. Generated captions call:

```http
POST /api/generate-captions
```

Backend checks before Gemini is called:

- valid Supabase session token in production
- existing or newly-created profile row when signed in
- refreshed monthly usage period when signed in
- remaining generated-caption credits when signed in
- request rate limit
- `GEMINI_API_KEY` presence

For local development, `ALLOW_UNAUTHENTICATED_CAPTION_TESTING="true"` lets the backend call Gemini without a Supabase session. Set it to `"false"` in production.

Caption credits are reserved atomically before Gemini is called through the `reserve_caption_credits()` database function. If Gemini fails, the backend refunds credits through `refund_caption_credits()` and returns fallback captions when possible:

- `Figure: Add description`
- `Table: Add description`

Only minimal asset context is sent: image data for figures and limited table rows/cells. The full document is not sent.

## Privacy Behavior

- Uploaded files are processed in the browser for parsing/export.
- Signed-in uploads are temporarily synced to the private Supabase bucket for account traceability, then deleted by default.
- Users can manually edit captions without consuming caption credits.
- Generated captions use only minimal figure/table context.
- Supabase stores profile, usage, document metadata, and temporary storage paths. Full files are not retained unless `VITE_PERSIST_UPLOADED_DOCS="true"`.

## Local Setup

Install dependencies:

```bash
npm install
```

Create `.env` from `.env.example` and set:

```bash
CAPTION_PROVIDER="gemini"
GEMINI_API_KEY="your_backend_only_key"
GEMINI_MODEL="gemini-2.0-flash"
VITE_SUPABASE_URL="https://your-project.supabase.co"
VITE_SUPABASE_ANON_KEY="your_public_anon_key"
SUPABASE_URL="https://your-project.supabase.co"
SUPABASE_SERVICE_ROLE_KEY="your_service_role_key"
PORT=8787
ALLOW_UNAUTHENTICATED_CAPTION_TESTING="true"
VITE_PERSIST_UPLOADED_DOCS="false"
```

Run the backend API:

```bash
npm run dev:api
```

Run the Vite app:

```bash
npm run dev
```

The Vite dev server proxies `/api` to `http://localhost:8787`.

Protected backend endpoints:

- `GET /api/account`
- `POST /api/process-document`
- `POST /api/export-document`
- `POST /api/generate-captions`

## Testing

To test profile creation:

1. Start the backend and frontend.
2. Sign in with Google or email.
3. Open Supabase Table Editor and confirm a row appears in `profiles`.
4. Confirm the account page shows free access and remaining usage.

To reset a test user:

```sql
update public.profiles
set plan = 'free',
    documents_used_this_month = 0,
    exports_used_this_month = 0,
    caption_credits_remaining = 5000,
    billing_period_start = now(),
    billing_period_end = now() + interval '1 month',
    updated_at = now()
where email = 'your-test-email@example.com';
```

## Build And Checks

```bash
npm run lint
npm run test:parser
npm run test:docx
npm run build
```

`test:parser` builds an in-memory `.docx` and verifies DrawingML images, VML images, table detection, previews, and audit counts.

`test:docx` builds an in-memory `.docx`, exports it through the real exporter, then inspects `word/document.xml` and `word/settings.xml` for Word-native TOC/list fields, `SEQ Figure` and `SEQ Table` captions, local file path cleanup, invalid page-number start cleanup, and `updateFields` settings.

## Deployment Notes

- Deploy the frontend as a static Vite app.
- Deploy `server/index.ts` as a Node/Express service or convert it to your platform's serverless function format.
- For GitHub Pages, set `VITE_BASE_PATH` to `"/your-repo-name/"` and `VITE_API_BASE_URL` to your hosted backend URL.
- Store Gemini and Supabase service role secrets only in backend environment variables.
- Never expose `SUPABASE_SERVICE_ROLE_KEY` or `GEMINI_API_KEY` in Vite variables.
- Set `CLIENT_ORIGIN`, `PRODUCTION_ORIGIN`, or `APP_URL` to your deployed frontend origin so CORS accepts production traffic.
- See [docs/deployment.md](</C:/Users/User/Downloads/doccaptioner/docs/deployment.md>) for the folder layout and GitHub Pages workflow.

## Limitations

- Word field numbering and generated lists are finalized when Microsoft Word updates fields after opening the exported file.
- Generated captions depend on the configured backend caption provider and its availability.
- Very large image-heavy documents may need smaller batches for caption generation.
- Cross-device visual QA should still be run in a browser before production release.
