# Deployment

DocuCaption Pro has two deployable parts:

- Static frontend: Vite React app from `dist/`
- Backend API: Express service in `server/index.ts`

GitHub Pages can host the frontend only. The backend must be hosted on a Node-capable service such as Render, Railway, Fly.io, a VPS, or serverless functions adapted from the Express routes.

## Folder Layout

- `src/`: frontend source
- `server/`: backend API and Gemini/Supabase service-role logic
- `public/`: favicon and web app manifest copied into the build
- `scripts/`: local regression tests
- `supabase/`: SQL schema and policies
- `docs/`: deployment and operations notes
- `dist/`: generated frontend build, ignored by git
- `.tools/`: local tooling/cache, ignored by git

## Frontend Environment

Only expose browser-safe variables:

```bash
VITE_SUPABASE_URL="https://your-project.supabase.co"
VITE_SUPABASE_ANON_KEY="your_public_anon_key"
VITE_API_BASE_URL="https://your-backend.example.com"
VITE_BASE_PATH="/your-repo-name/"
```

Use `VITE_BASE_PATH="/"` for a custom domain or root deployment.

Never expose:

```bash
GEMINI_API_KEY
SUPABASE_SERVICE_ROLE_KEY
```

## Backend Environment

Set these only on the backend host:

```bash
CAPTION_PROVIDER="gemini"
GEMINI_API_KEY="your_backend_only_key"
GEMINI_MODEL="gemini-2.5-flash"
SUPABASE_URL="https://your-project.supabase.co"
SUPABASE_SERVICE_ROLE_KEY="your_service_role_key"
CLIENT_ORIGIN="https://your-github-username.github.io"
PRODUCTION_ORIGIN="https://your-github-username.github.io"
PORT=8787
ALLOW_UNAUTHENTICATED_CAPTION_TESTING="false"
```

If your GitHub Pages URL is a project site, set the origin without the path:

```bash
CLIENT_ORIGIN="https://your-github-username.github.io"
```

## Build Locally

```bash
npm install
npm run lint
npm run test:parser
npm run test:docx
npm run build
```

For a GitHub Pages project site:

```bash
$env:VITE_BASE_PATH="/your-repo-name/"
$env:VITE_API_BASE_URL="https://your-backend.example.com"
npm run build
```

## GitHub Pages

The workflow in `.github/workflows/deploy-pages.yml` builds and publishes `dist/`.

Repository settings:

1. Go to Settings -> Pages.
2. Set Source to GitHub Actions.
3. Add repository variables:
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_ANON_KEY`
   - `VITE_API_BASE_URL`
   - `VITE_BASE_PATH`
4. Push to `main`.

For `VITE_BASE_PATH`, use:

```text
/repository-name/
```

For a custom domain, use:

```text
/
```

## Backend Hosting

The backend start command is:

```bash
npm run dev:api
```

For production, use a process manager or platform command that runs:

```bash
npx tsx server/index.ts
```

Set CORS env vars to your frontend origin. The frontend calls the backend using `VITE_API_BASE_URL`.

## Supabase

Run `supabase/schema.sql` in Supabase SQL editor before production use.

Required setup:

- Enable Email provider.
- Enable Google provider if you want Google sign-in.
- Add your GitHub Pages URL to Supabase Auth redirect URLs.
- Confirm RLS policies are enabled.

## Files Not To Deploy

These are ignored and should not be committed or deployed:

- `.env`
- `.env.local`
- `.tools/`
- `node_modules/`
- `dist/`
- `*.log`
