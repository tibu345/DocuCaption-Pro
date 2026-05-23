# DocuCaption Pro

DocuCaption Pro helps clean up Word reports before submission. Upload a `.docx`, detect figures and tables, generate editable captions, prevent duplicate captions, and export a corrected Word document.

Live app:

https://tibu345.github.io/DocuCaption-Pro/

## What It Does

- Detects figures and tables in Word `.docx` files
- Finds missing or duplicate captions
- Generates draft captions for figures and tables
- Lets you edit captions before export
- Adds Word-native `Figure` and `Table` caption fields
- Builds or updates the Table of Contents, Table of Figures, and Table of Tables
- Exports a corrected `.docx` copy without changing your original file

## How To Use It

1. Open the live app.
2. Sign in with email or Google.
3. Upload a `.docx` report.
4. Review the detected figures and tables.
5. Generate captions or edit them manually.
6. Export the corrected Word file.
7. Open the exported file in Microsoft Word and update fields if Word asks.

## Privacy Notes

- The app works with `.docx` files only.
- Parsing and export are handled in the browser.
- Generated captions use only figure/table context, not your full document text.
- Uploaded files are not meant to be permanently stored unless storage persistence is explicitly enabled by the app owner.

## Current Limits

The app is free to use, with usage limits to protect the caption generation backend:

- 1000 document scans per month
- 1000 exports per month
- 5000 generated captions per month
- 200 figures/tables per document

## Tech Stack

- React
- Vite
- TypeScript
- Supabase Auth and database
- Express backend API
- Gemini caption generation
- GitHub Pages frontend hosting
- Render backend hosting

## Local Development

Install dependencies:

```bash
npm install
```

Create `.env` from `.env.example`, then set your local Supabase and Gemini values.

Run the backend:

```bash
npm run dev:api
```

Run the frontend:

```bash
npm run dev
```

Local URLs:

```txt
Frontend: http://localhost:3000
Backend:  http://localhost:8787
```

## Required Environment Variables

Frontend-safe variables:

```txt
VITE_SUPABASE_URL
VITE_SUPABASE_ANON_KEY
VITE_API_BASE_URL
VITE_BASE_PATH
```

Backend-only secrets:

```txt
SUPABASE_URL
SUPABASE_SERVICE_ROLE_KEY
GEMINI_API_KEY
GEMINI_MODEL
CAPTION_PROVIDER
```

Never expose backend-only secrets in frontend hosting variables.

## Deployment

Frontend:

- Hosted on GitHub Pages
- Build command: `npm run build`
- Output folder: `dist`
- Base path: `/DocuCaption-Pro/`

Backend:

- Hosted as a Node web service
- Start command: `npm start`
- API file: `server/index.ts`

Database/Auth:

- Supabase project
- Run `supabase/schema.sql` in the Supabase SQL editor
- Add the live app URL to Supabase Auth redirect URLs:

```txt
https://tibu345.github.io/DocuCaption-Pro/
```

## Checks

```bash
npm run lint
npm run test:parser
npm run test:docx
npm run build
```

## Repository Structure

```txt
src/        frontend app
server/     backend API
supabase/   database schema
scripts/    regression checks
docs/       deployment notes
public/     static assets
```

## Notes

- Word field numbering is finalized by Microsoft Word when fields are updated.
- Generated captions depend on the backend caption provider being available.
- Very large image-heavy documents may take longer to process.
