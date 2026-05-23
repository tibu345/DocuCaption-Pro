import dotenv from 'dotenv';
import express from 'express';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createClient, type SupabaseClient, type User as SupabaseUser } from '@supabase/supabase-js';
import { ACCESS_PLANS, normalizePlanId } from '../src/lib/access';
import { sanitizeCaptionDescription } from '../src/lib/captionUtils';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const envPath = path.join(projectRoot, '.env');
const envLocalPath = path.join(projectRoot, '.env.local');
const envResult = fs.existsSync(envPath) ? dotenv.config({ path: envPath }) : undefined;
const envLocalResult = fs.existsSync(envLocalPath) ? dotenv.config({ path: envLocalPath, override: true }) : undefined;
const loadedEnvFiles = [
  envResult?.parsed ? '.env' : undefined,
  envLocalResult?.parsed ? '.env.local' : undefined,
].filter((file): file is string => Boolean(file));

interface CaptionRequestElement {
  id: string;
  type: 'image' | 'table';
  alt?: string;
  imageDataUrl?: string;
  rows?: string[][];
}

interface CaptionRequest {
  elements?: CaptionRequestElement[];
}

interface DocumentStorageRequest {
  storageBucket?: string;
  storagePath?: string;
  deleteAfterProcessing?: boolean;
}

interface ProfileRow {
  id: string;
  email: string;
  full_name: string | null;
  plan: string;
  subscription_status: 'inactive' | 'active' | 'past_due' | 'canceled';
  documents_used_this_month: number;
  exports_used_this_month: number;
  caption_credits_remaining: number;
  billing_period_start: string;
  billing_period_end: string;
  created_at?: string;
  updated_at?: string;
}

interface AccountResponse {
  profile: ProfileRow;
  usage: {
    documentsRemaining: number;
    exportsRemaining: number;
    captionCreditsRemaining: number;
    billingPeriodEnd: string;
  };
}

const app = express();
const port = Number(process.env.PORT ?? 8787);
const host = process.env.HOST || '0.0.0.0';
const requests = new Map<string, { count: number; resetAt: number }>();
const supabaseAdmin = createSupabaseAdminClient();

app.use((req, res, next) => {
  const origin = req.header('Origin');
  const allowedOrigins = getAllowedOrigins();
  if (origin && allowedOrigins.has(origin)) {
    res.header('Access-Control-Allow-Origin', origin);
    res.header('Vary', 'Origin');
  }
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  if (req.method === 'OPTIONS') {
    res.sendStatus(204);
    return;
  }
  next();
});

app.use(express.json({ limit: process.env.JSON_BODY_LIMIT || '12mb' }));

app.use((error: unknown, _req: express.Request, res: express.Response, next: express.NextFunction) => {
  if (error instanceof Error && 'type' in error && error.type === 'entity.too.large') {
    res.status(413).json({
      error: 'Caption request is too large. Try fewer figures at once or use smaller images.',
    });
    return;
  }
  next(error);
});

app.get('/api/account', async (req, res) => {
  try {
    const context = await getRequestContext(req);
    res.json(createAccountResponse(context.profile));
  } catch (error) {
    sendApiError(res, error);
  }
});
app.get('/api/me', async (req, res) => {
  try {
    const context = await getRequestContext(req);
    res.json(createAccountResponse(context.profile));
  } catch (error) {
    sendApiError(res, error);
  }
});

app.post('/api/process-document', processDocumentHandler);
app.post('/api/documents/consume', processDocumentHandler);

async function processDocumentHandler(req: express.Request, res: express.Response) {
  try {
    const context = await getRequestContext(req);
    const { fileName, assetCount } = req.body as { fileName?: string; assetCount?: number };
    const storage = normalizeDocumentStorageRequest(context.user.id, req.body as DocumentStorageRequest);
    const plan = getProfilePlan(context.profile);
    const assets = Number(assetCount ?? 0);

    if (assets > plan.maxAssetsPerDocument) {
      await insertUsageLog(context.user.id, 'caption_denied', 0, { reason: 'asset_limit', assetCount: assets, plan: plan.id });
      throw apiError(403, `${plan.name} supports up to ${plan.maxAssetsPerDocument} figures/tables per document.`);
    }

    const profile = await consumeDocumentQuota(context.user.id, plan.documentLimit);
    if (!profile) {
      await insertUsageLog(context.user.id, 'caption_denied', 0, { reason: 'document_limit', fileName, plan: plan.id });
      throw apiError(403, 'You have reached your monthly document limit.');
    }

    await insertDocumentLog(context.user.id, fileName ?? 'document.docx', assets, 'processed', storage);
    await insertUsageLog(context.user.id, 'document_processed', 1, {
      fileName,
      assetCount: assets,
      plan: plan.id,
      storedTemporarily: Boolean(storage.storage_path),
    });
    res.json(createAccountResponse(profile));
  } catch (error) {
    sendApiError(res, error);
  }
}

app.post('/api/export-document', exportDocumentHandler);
app.post('/api/exports/consume', exportDocumentHandler);

async function exportDocumentHandler(req: express.Request, res: express.Response) {
  try {
    const context = await getRequestContext(req);
    const { fileName } = req.body as { fileName?: string };
    const plan = getProfilePlan(context.profile);

    const profile = await consumeExportQuota(context.user.id, plan.exportLimit);
    if (!profile) {
      await insertUsageLog(context.user.id, 'caption_denied', 0, { reason: 'export_limit', fileName, plan: plan.id });
      throw apiError(403, 'You have reached your monthly export limit.');
    }

    await insertUsageLog(context.user.id, 'export_created', 1, { fileName, plan: plan.id });
    res.json(createAccountResponse(profile));
  } catch (error) {
    sendApiError(res, error);
  }
}

app.post('/api/generate-captions', generateCaptionsHandler);
app.post('/api/captions', generateCaptionsHandler);

async function generateCaptionsHandler(req: express.Request, res: express.Response) {
  const body = req.body as CaptionRequest;
  let reservedCredits = 0;
  let contextForRefund: { user: SupabaseUser; profile: ProfileRow } | undefined;
  try {
    const context = await getOptionalRequestContext(req);
    const userKey = context?.user.id ?? req.ip ?? 'local-caption-test';
    contextForRefund = context;
    const plan = context ? getProfilePlan(context.profile) : ACCESS_PLANS.free;
    const elements = body.elements ?? [];
    const creditsNeeded = elements.length;

    if (!allowRequest(userKey)) {
      throw apiError(429, 'Rate limit exceeded. Please wait before requesting more generated captions.');
    }
    if (!context && !allowUnauthenticatedCaptionTesting()) {
      throw apiError(401, 'Sign in before using generated captions.');
    }
    if (context && !plan.generatedCaptionsEnabled) {
      await insertUsageLog(context.user.id, 'caption_denied', 0, { reason: 'plan', plan: plan.id, captionsRequested: creditsNeeded });
      throw apiError(403, 'Generated captions are unavailable.');
    }
    if (context) {
      const reservedProfile = await reserveCaptionCredits(context.user.id, creditsNeeded);
      if (!reservedProfile) {
        await insertUsageLog(context.user.id, 'caption_denied', 0, {
          reason: 'credits',
          captionsRequested: creditsNeeded,
          plan: plan.id,
        });
        throw apiError(402, 'You have no caption credits remaining.');
      }
      reservedCredits = creditsNeeded;
    }
    const provider = getCaptionProvider();
    if (provider === 'gemini' && !isConfiguredSecret(process.env.GEMINI_API_KEY)) {
      throw apiError(503, 'Gemini key is not configured on the backend.');
    }

    const captions = provider === 'ollama' ? await generateOllamaCaptions(elements) : await generateGeminiCaptions(elements);
    const creditsUsed = Math.min(captions.length, creditsNeeded);
    if (context && reservedCredits > creditsUsed) {
      await refundCaptionCredits(context.user.id, reservedCredits - creditsUsed);
    }
    const profile = context ? await fetchProfile(context.user.id) : undefined;
    if (context) {
      await insertUsageLog(context.user.id, 'captions_generated', creditsUsed, {
        captionsRequested: creditsNeeded,
        plan: plan.id,
      });
    }
    res.json({ captions, creditsUsed: context ? creditsUsed : 0, profile });
  } catch (error) {
    if (isApiError(error)) {
      if (reservedCredits > 0 && contextForRefund) {
        await refundCaptionCredits(contextForRefund.user.id, reservedCredits).catch(console.error);
      }
      sendApiError(res, error);
      return;
    }
    console.error(error);
    const detail = error instanceof Error ? error.message : 'Unknown Gemini API error.';
    if (reservedCredits > 0 && contextForRefund) {
      await refundCaptionCredits(contextForRefund.user.id, reservedCredits).catch(console.error);
    }
    if (isRecoverableAiFailure(detail)) {
      const context = await getRequestContext(req).catch(() => undefined);
      if (context) await insertUsageLog(context.user.id, 'captions_fallback', 0, { reason: detail, captionsRequested: body.elements?.length ?? 0, plan: context.profile.plan });
      res.json({
        captions: (body.elements ?? []).map(fallbackCaptionFor),
        creditsUsed: 0,
        profile: context?.profile,
        warning: createUserFacingCaptionWarning(detail),
      });
      return;
    }
    res.status(502).send(`Generated caption service failed: ${detail}`);
  }
}

app.get('/api/admin/usage', (_req, res) => {
  res.json({
    users: [],
    note: 'Connect a database to list users, reset credits, change plans, and inspect failed processing jobs.',
  });
});

app.get('/api/health', (_req, res) => {
  const provider = getCaptionProvider();
  res.json({
    ok: true,
    captionProvider: provider,
    geminiConfigured: isConfiguredSecret(process.env.GEMINI_API_KEY),
    geminiModel: process.env.GEMINI_MODEL || 'gemini-2.0-flash',
    supabaseServerConfigured: Boolean(supabaseAdmin),
    ollamaConfigured: Boolean(process.env.OLLAMA_BASE_URL && process.env.OLLAMA_MODEL),
    ollamaBaseUrl: process.env.OLLAMA_BASE_URL || 'http://localhost:11434',
    ollamaModel: process.env.OLLAMA_MODEL || 'llama3.2:1b',
    envFilesLoaded: loadedEnvFiles,
  });
});

app.listen(port, host, () => {
  console.log(`DocuCaption API listening on http://${host}:${port}`);
});

function createSupabaseAdminClient(): SupabaseClient | undefined {
  const supabaseUrl = process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL;
  const serviceRoleKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!supabaseUrl || !serviceRoleKey || /your-project|your_service_role/i.test(`${supabaseUrl} ${serviceRoleKey}`)) return undefined;
  if (!isValidHttpUrl(supabaseUrl)) {
    console.warn('Ignoring invalid SUPABASE_URL. It must be a full http(s) URL, for example https://project-ref.supabase.co.');
    return undefined;
  }
  return createClient(supabaseUrl, serviceRoleKey, {
    auth: {
      persistSession: false,
      autoRefreshToken: false,
    },
  });
}

function isValidHttpUrl(value: string): boolean {
  try {
    const url = new URL(value);
    return url.protocol === 'http:' || url.protocol === 'https:';
  } catch {
    return false;
  }
}

async function getRequestContext(req: express.Request): Promise<{ user: SupabaseUser; profile: ProfileRow }> {
  if (!supabaseAdmin) {
    throw apiError(503, 'Supabase server configuration is missing. Add SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY.');
  }

  const token = req.header('Authorization')?.replace(/^Bearer\s+/i, '').trim();
  if (!token) throw apiError(401, 'Missing authentication token.');

  const { data, error } = await supabaseAdmin.auth.getUser(token);
  if (error || !data.user) throw apiError(401, 'Invalid or expired authentication token.');

  await ensureProfile(data.user);
  return { user: data.user, profile: await refreshUserBillingPeriod(data.user.id) };
}

async function getOptionalRequestContext(req: express.Request): Promise<{ user: SupabaseUser; profile: ProfileRow } | undefined> {
  const authorization = req.header('Authorization');
  if (!authorization) return undefined;
  return getRequestContext(req);
}

function createAccountResponse(profile: ProfileRow): AccountResponse {
  const plan = getProfilePlan(profile);
  return {
    profile,
    usage: {
      documentsRemaining: Math.max(0, plan.documentLimit - profile.documents_used_this_month),
      exportsRemaining: Math.max(0, plan.exportLimit - profile.exports_used_this_month),
      captionCreditsRemaining: profile.caption_credits_remaining,
      billingPeriodEnd: profile.billing_period_end,
    },
  };
}

function getProfilePlan(profile: ProfileRow) {
  return ACCESS_PLANS[normalizePlanId(profile.plan)];
}

async function ensureProfile(user: SupabaseUser): Promise<ProfileRow> {
  if (!supabaseAdmin) throw apiError(503, 'Supabase server configuration is missing.');
  const { data, error } = await supabaseAdmin.from('profiles').select('*').eq('id', user.id).maybeSingle<ProfileRow>();
  if (error) throw apiError(500, `Could not fetch profile: ${error.message}`);
  if (data) return data;

  const now = new Date();
  const periodEnd = new Date(now);
  periodEnd.setUTCMonth(periodEnd.getUTCMonth() + 1);
  const fullName = typeof user.user_metadata?.full_name === 'string' ? user.user_metadata.full_name : null;
  const insert: ProfileRow = {
    id: user.id,
    email: user.email ?? '',
    full_name: fullName,
    plan: 'free',
    subscription_status: 'inactive',
    documents_used_this_month: 0,
    exports_used_this_month: 0,
    caption_credits_remaining: ACCESS_PLANS.free.captionCredits,
    billing_period_start: now.toISOString(),
    billing_period_end: periodEnd.toISOString(),
  };
  const { data: created, error: createError } = await supabaseAdmin.from('profiles').insert(insert).select('*').single<ProfileRow>();
  if (createError || !created) throw apiError(500, `Could not create profile: ${createError?.message ?? 'unknown error'}`);
  return created;
}

async function fetchProfile(userId: string): Promise<ProfileRow> {
  if (!supabaseAdmin) throw apiError(503, 'Supabase server configuration is missing.');
  const { data, error } = await supabaseAdmin.from('profiles').select('*').eq('id', userId).single<ProfileRow>();
  if (error || !data) throw apiError(500, `Could not fetch profile: ${error?.message ?? 'unknown error'}`);
  return data;
}

async function refreshUserBillingPeriod(userId: string): Promise<ProfileRow> {
  return resetBillingPeriodIfNeeded(await fetchProfile(userId));
}

async function resetBillingPeriodIfNeeded(profile: ProfileRow): Promise<ProfileRow> {
  if (new Date(profile.billing_period_end).getTime() > Date.now()) return profile;
  const plan = getProfilePlan(profile);
  const start = new Date();
  const end = new Date(start);
  end.setUTCMonth(end.getUTCMonth() + 1);
  return updateProfileUsage(profile.id, {
    documents_used_this_month: 0,
    exports_used_this_month: 0,
    caption_credits_remaining: plan.captionCredits,
    billing_period_start: start.toISOString(),
    billing_period_end: end.toISOString(),
  });
}

async function reserveCaptionCredits(userId: string, amount: number): Promise<ProfileRow | undefined> {
  if (!supabaseAdmin) throw apiError(503, 'Supabase server configuration is missing.');
  if (amount <= 0) return fetchProfile(userId);
  const { data, error } = await supabaseAdmin.rpc('reserve_caption_credits', {
    target_user_id: userId,
    credit_amount: amount,
  });
  if (error) throw apiError(500, `Could not reserve caption credits: ${error.message}`);
  return firstRpcProfile(data);
}

async function refundCaptionCredits(userId: string, amount: number): Promise<ProfileRow | undefined> {
  if (!supabaseAdmin) throw apiError(503, 'Supabase server configuration is missing.');
  if (amount <= 0) return fetchProfile(userId);
  const { data, error } = await supabaseAdmin.rpc('refund_caption_credits', {
    target_user_id: userId,
    credit_amount: amount,
  });
  if (error) throw apiError(500, `Could not refund caption credits: ${error.message}`);
  return firstRpcProfile(data);
}

async function consumeDocumentQuota(userId: string, monthlyLimit: number): Promise<ProfileRow | undefined> {
  if (!supabaseAdmin) throw apiError(503, 'Supabase server configuration is missing.');
  const { data, error } = await supabaseAdmin.rpc('consume_document_quota', {
    target_user_id: userId,
    monthly_limit: monthlyLimit,
  });
  if (error) throw apiError(500, `Could not consume document quota: ${error.message}`);
  return firstRpcProfile(data);
}

async function consumeExportQuota(userId: string, monthlyLimit: number): Promise<ProfileRow | undefined> {
  if (!supabaseAdmin) throw apiError(503, 'Supabase server configuration is missing.');
  const { data, error } = await supabaseAdmin.rpc('consume_export_quota', {
    target_user_id: userId,
    monthly_limit: monthlyLimit,
  });
  if (error) throw apiError(500, `Could not consume export quota: ${error.message}`);
  return firstRpcProfile(data);
}

function firstRpcProfile(data: unknown): ProfileRow | undefined {
  return Array.isArray(data) && data.length > 0 ? (data[0] as ProfileRow) : undefined;
}

async function updateProfileUsage(userId: string, patch: Partial<ProfileRow>): Promise<ProfileRow> {
  if (!supabaseAdmin) throw apiError(503, 'Supabase server configuration is missing.');
  const { data, error } = await supabaseAdmin
    .from('profiles')
    .update({ ...patch, updated_at: new Date().toISOString() })
    .eq('id', userId)
    .select('*')
    .single<ProfileRow>();
  if (error || !data) throw apiError(500, `Could not update profile: ${error?.message ?? 'unknown error'}`);
  return data;
}

async function insertDocumentLog(
  userId: string,
  fileName: string,
  assetCount: number,
  status: string,
  storage?: { storage_bucket?: string; storage_path?: string; deleted_at?: string },
): Promise<void> {
  if (!supabaseAdmin) return;
  const { error } = await supabaseAdmin.from('documents').insert({
    user_id: userId,
    file_name: fileName,
    asset_count: assetCount,
    status,
    ...storage,
  });
  if (error) throw apiError(500, `Could not record document usage: ${error.message}`);
}

async function insertUsageLog(userId: string, action: string, amount: number, metadata: Record<string, unknown>): Promise<void> {
  if (!supabaseAdmin) return;
  const { error } = await supabaseAdmin.from('usage_logs').insert({
    user_id: userId,
    action,
    amount,
    metadata,
  });
  if (error) throw apiError(500, `Could not record usage: ${error.message}`);
}

function apiError(status: number, message: string): Error & { status: number } {
  return Object.assign(new Error(message), { status });
}

function isApiError(error: unknown): error is Error & { status: number } {
  return error instanceof Error && typeof (error as { status?: unknown }).status === 'number';
}

function sendApiError(res: express.Response, error: unknown): void {
  const status = isApiError(error) ? error.status : 500;
  const message = createClientSafeErrorMessage(error instanceof Error ? error.message : 'Unexpected server error.');
  res.status(status).json({ error: message });
}

function createClientSafeErrorMessage(message: string): string {
  if (/jwt|authentication token|auth session/i.test(message)) return 'Your session expired. Please sign in again.';
  if (/service_role|supabase server configuration/i.test(message)) return 'The document service is not fully configured.';
  if (/reserve caption credits|refund caption credits|consume document quota|consume export quota/i.test(message)) {
    return 'Usage could not be updated. Please try again.';
  }
  if (/record document usage|record usage|create profile|fetch profile|update profile/i.test(message)) {
    return 'Account usage could not be synced. Please try again.';
  }
  if (/gemini|api key|permission|quota|resource_exhausted|billing/i.test(message)) {
    return 'Generated captions are temporarily unavailable.';
  }
  return message;
}

function normalizeDocumentStorageRequest(
  userId: string,
  request: DocumentStorageRequest,
): { storage_bucket?: string; storage_path?: string; deleted_at?: string } {
  if (!request.storagePath) return {};
  const bucket = request.storageBucket || 'docucaption-documents';
  if (bucket !== 'docucaption-documents') throw apiError(400, 'Unsupported document storage bucket.');
  if (!request.storagePath.startsWith(`${userId}/`)) throw apiError(403, 'Document storage path does not match the signed-in user.');
  return {
    storage_bucket: bucket,
    storage_path: request.storagePath,
    deleted_at: request.deleteAfterProcessing ? new Date().toISOString() : undefined,
  };
}

function getAllowedOrigins(): Set<string> {
  return new Set(
    [
      'http://localhost:3000',
      'http://127.0.0.1:3000',
      process.env.APP_URL,
      process.env.CLIENT_ORIGIN,
      process.env.PRODUCTION_ORIGIN,
    ]
      .filter((origin): origin is string => Boolean(origin))
      .map((origin) => origin.replace(/\/+$/g, '')),
  );
}

function allowUnauthenticatedCaptionTesting(): boolean {
  return process.env.ALLOW_UNAUTHENTICATED_CAPTION_TESTING !== 'false' && process.env.NODE_ENV !== 'production';
}

function allowRequest(userId: string): boolean {
  const now = Date.now();
  const windowMs = 60_000;
  const maxRequests = 12;
  const existing = requests.get(userId);
  if (!existing || existing.resetAt <= now) {
    requests.set(userId, { count: 1, resetAt: now + windowMs });
    return true;
  }
  if (existing.count >= maxRequests) return false;
  existing.count += 1;
  return true;
}

async function generateGeminiCaptions(elements: CaptionRequestElement[]): Promise<{ id: string; caption: string }[]> {
  const minimalContext = elements.map((element) => ({
    id: element.id,
    type: element.type,
    label: labelForElement(element),
    alt: cleanFallbackText(element.alt).slice(0, 240),
    hasImageData: Boolean(parseDataUrlImage(element.imageDataUrl)),
    rows: element.rows?.slice(0, 5).map((row) => row.map((cell) => cell.slice(0, 120))),
  }));
  const model = process.env.GEMINI_MODEL || 'gemini-2.0-flash';
  const parts = createGeminiParts(elements, minimalContext);

  const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${process.env.GEMINI_API_KEY}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: [
        {
          role: 'user',
          parts,
        },
      ],
      generationConfig: {
        responseMimeType: 'application/json',
        responseSchema: {
          type: 'ARRAY',
          items: {
            type: 'OBJECT',
            properties: {
              id: { type: 'STRING' },
              caption: { type: 'STRING' },
            },
            required: ['id', 'caption'],
          },
        },
      },
    }),
  });

  if (!response.ok) {
    throw new Error(summarizeGeminiError(await response.text()));
  }

  const payload = (await response.json()) as { candidates?: { content?: { parts?: { text?: string }[] } }[] };
  const text = payload.candidates?.[0]?.content?.parts?.[0]?.text ?? '[]';
  const parsed = parseCaptionJson(text);
  const captions = parsed.map((caption) => ({
    id: caption.id,
    caption: caption.caption?.trim() || 'Add description',
  }));
  return elements.map((element) => normalizeNumberedCaption(element, captions.find((caption) => caption.id === element.id) ?? fallbackCaptionFor(element)));
}

async function generateOllamaCaptions(elements: CaptionRequestElement[]): Promise<{ id: string; caption: string }[]> {
  const baseUrl = (process.env.OLLAMA_BASE_URL || 'http://localhost:11434').replace(/\/+$/g, '');
  const model = process.env.OLLAMA_MODEL || 'llama3.2:1b';
  const minimalContext = elements.map((element) => ({
    id: element.id,
    type: element.type,
    label: labelForElement(element),
    alt: cleanFallbackText(element.alt).slice(0, 240),
    hasImageData: Boolean(parseDataUrlImage(element.imageDataUrl)),
    rows: element.rows?.slice(0, 5).map((row) => row.map((cell) => cell.slice(0, 120))),
  }));

  const response = await fetch(`${baseUrl}/api/generate`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      model,
      stream: false,
      format: 'json',
      options: {
        temperature: 0.2,
        num_predict: 700,
      },
      prompt: [
        'You generate concise captions for academic Word documents.',
        'Use only the provided context. Do not invent entities, values, trends, or conclusions.',
        'Return ONLY valid JSON. Do not use markdown.',
        'The JSON must be an array of objects with exactly these keys: id, caption.',
        'Use the provided label at the start of each caption, e.g. "Figure 1:" or "Table 1:".',
        'If context is weak, use "Figure: Add description" or "Table: Add description".',
        `Assets: ${JSON.stringify(minimalContext)}`,
      ].join('\n'),
    }),
  });

  if (!response.ok) {
    throw new Error(`Ollama request failed: ${await response.text()}`);
  }

  const payload = (await response.json()) as { response?: string; error?: string };
  if (payload.error) throw new Error(`Ollama error: ${payload.error}`);
  const parsed = parseCaptionJson(payload.response ?? '[]');
  const captions = parsed.map((caption) => ({
    id: caption.id,
    caption: caption.caption?.trim() || 'Add description',
  }));
  return elements.map((element) => normalizeNumberedCaption(element, captions.find((caption) => caption.id === element.id) ?? fallbackCaptionFor(element)));
}

function summarizeGeminiError(raw: string): string {
  try {
    const parsed = JSON.parse(raw) as { error?: { message?: string; status?: string; code?: number } };
    const message = parsed.error?.message ?? raw;
    const status = parsed.error?.status ? ` (${parsed.error.status})` : '';
    return `${message}${status}`;
  } catch {
    return raw;
  }
}

function parseCaptionJson(text: string): { id: string; caption: string }[] {
  const trimmed = text.trim();
  const withoutFence = trimmed.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '');
  const parsed = JSON.parse(withoutFence) as unknown;
  if (!Array.isArray(parsed)) return [];
  return parsed
    .filter((item): item is { id: string; caption: string } => Boolean(item) && typeof item === 'object' && 'id' in item && 'caption' in item)
    .map((item) => ({ id: String(item.id), caption: String(item.caption) }));
}

function createGeminiParts(elements: CaptionRequestElement[], minimalContext: unknown[]): Array<{ text: string } | { inline_data: { mime_type: string; data: string } }> {
  const parts: Array<{ text: string } | { inline_data: { mime_type: string; data: string } }> = [
    {
      text: [
        'You generate concise captions for academic Word documents.',
        'Ground every figure caption in the visible image content supplied after its asset ID.',
        'Ground every table caption in the supplied table rows.',
        'Do not invent names, values, trends, labels, or conclusions that are not visible/provided.',
        'If the image is unclear or context is insufficient, use "Figure: Add description" or "Table: Add description".',
        'Use the provided label at the start of each caption, e.g. "Figure 1:" or "Table 1:".',
        'Return ONLY valid JSON. Do not use markdown.',
        'The JSON must be an array of objects with exactly these keys: id, caption.',
        `Asset metadata: ${JSON.stringify(minimalContext)}`,
      ].join('\n'),
    },
  ];

  for (const element of elements) {
    if (element.type !== 'image') continue;
    const image = parseDataUrlImage(element.imageDataUrl);
    if (!image) continue;
    parts.push({ text: `Image asset ${element.id}. Use this image only for caption ${element.id}.` });
    parts.push({ inline_data: { mime_type: image.mimeType, data: image.data } });
  }

  return parts;
}

function parseDataUrlImage(dataUrl: string | undefined): { mimeType: string; data: string } | undefined {
  if (!dataUrl) return undefined;
  const match = dataUrl.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (!match) return undefined;
  return {
    mimeType: match[1],
    data: match[2],
  };
}

function fallbackCaptionFor(element: CaptionRequestElement): { id: string; caption: string } {
  const label = labelForElement(element);
  const description = element.type === 'image' ? cleanFallbackText(element.alt) : summarizeTableRows(element.rows);
  return {
    id: element.id,
    caption: `${label}: ${description || 'Add description'}`,
  };
}

function normalizeNumberedCaption(element: CaptionRequestElement, caption: { id: string; caption: string }): { id: string; caption: string } {
  const label = labelForElement(element);
  const withoutLabel = caption.caption
    .replace(/^(figure|fig\.|table)\s*\d*\s*[:.)-]?\s*/i, '')
    .trim();
  return {
    id: caption.id,
    caption: `${label}: ${sanitizeCaptionDescription(withoutLabel)}`,
  };
}

function labelForElement(element: CaptionRequestElement): string {
  const match = element.id.match(/^(fig|tab)-(\d+)$/);
  const number = match ? Number(match[2]) + 1 : 1;
  return `${element.type === 'image' ? 'Figure' : 'Table'} ${number}`;
}

function isRecoverableAiFailure(message: string): boolean {
  return /quota|rate.?limit|resource_exhausted|api key|permission|billing|unavailable|ollama|fetch failed|econnrefused/i.test(message);
}

function createUserFacingCaptionWarning(message: string): string {
  if (/ollama|fetch failed|econnrefused/i.test(message)) {
    return 'Generated captions are temporarily unavailable. Fallback captions were used and no caption credits were deducted.';
  }
  if (/quota|resource_exhausted|billing|rate.?limit/i.test(message)) {
    return 'Generated captions are temporarily unavailable because the provider quota was reached. Fallback captions were used and no caption credits were deducted.';
  }
  if (/api key|permission/i.test(message)) {
    return 'Generated captions are not available because the backend provider key is missing or not permitted. Fallback captions were used and no caption credits were deducted.';
  }
  return 'Generated captions are temporarily unavailable. Fallback captions were used and no caption credits were deducted.';
}

function getCaptionProvider(): 'gemini' | 'ollama' {
  const provider = process.env.CAPTION_PROVIDER ?? process.env.AI_PROVIDER;
  return provider?.toLowerCase() === 'ollama' ? 'ollama' : 'gemini';
}

function isConfiguredSecret(value: string | undefined): boolean {
  return Boolean(value && value.trim() && !/paste_your|replace_me|your_/i.test(value));
}

function cleanFallbackText(value: string | undefined): string {
  const cleaned = (value ?? '')
    .replace(/\s+/g, ' ')
    .replace(/^Figure\s*[:.\-\d\s]*/i, '')
    .trim()
    .slice(0, 160);
  const sanitized = sanitizeCaptionDescription(cleaned);
  return sanitized === 'Add description' ? '' : sanitized;
}

function summarizeTableRows(rows: string[][] | undefined): string {
  const firstUsefulRow = rows?.find((row) => row.some((cell) => cell.trim().length > 0));
  if (!firstUsefulRow) return '';
  const cells = firstUsefulRow.map(cleanFallbackText).filter(Boolean).slice(0, 4);
  return cells.length > 0 ? `Summary of ${cells.join(', ')}` : '';
}
