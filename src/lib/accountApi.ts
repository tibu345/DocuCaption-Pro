import { createSignedOutUser, saveUser, type UserProfile } from './auth';
import { normalizePlanId } from './access';
import { supabase } from './supabaseClient';
import type { StoredDocumentRef } from './documentStorage';
import { friendlyErrorMessage } from './errorMessages';
import { apiFetch } from './apiClient';

export interface ServerProfile {
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
}

export interface AccountResponse {
  profile: ServerProfile;
  usage?: {
    documentsRemaining: number;
    exportsRemaining: number;
    captionCreditsRemaining: number;
    billingPeriodEnd: string;
  };
}

export async function fetchAccountProfile(): Promise<UserProfile> {
  return requestAccount('/api/account', { method: 'GET' });
}

export async function recordDocumentProcessed(fileName: string, assetCount: number, storage?: StoredDocumentRef): Promise<UserProfile> {
  return requestAccount('/api/process-document', {
    method: 'POST',
    body: JSON.stringify({
      fileName,
      assetCount,
      storageBucket: storage?.bucket,
      storagePath: storage?.path,
      deleteAfterProcessing: storage?.deleteAfterProcessing,
    }),
  });
}

export async function recordExport(fileName: string): Promise<UserProfile> {
  return requestAccount('/api/export-document', {
    method: 'POST',
    body: JSON.stringify({ fileName }),
  });
}

export async function getAccessToken(): Promise<string | undefined> {
  if (!supabase) return undefined;
  const { data } = await supabase.auth.getSession();
  return data.session?.access_token;
}

export function mapServerProfile(profile: ServerProfile): UserProfile {
  const user: UserProfile = {
    id: profile.id,
    email: profile.email,
    displayName: profile.full_name ?? profile.email.split('@')[0] ?? 'User',
    planId: normalizePlanId(profile.plan),
    paymentStatus: profile.subscription_status,
    authProvider: 'email',
    isAdmin: false,
    usage: {
      documentsProcessed: profile.documents_used_this_month,
      exportsUsed: profile.exports_used_this_month,
      captionCreditsRemaining: profile.caption_credits_remaining,
      failedProcessingJobs: 0,
      periodStartedAt: profile.billing_period_start,
      billingPeriodEnd: profile.billing_period_end,
    },
  };
  saveUser(user);
  return user;
}

async function requestAccount(path: string, init: RequestInit): Promise<UserProfile> {
  const token = await getAccessToken();
  if (!token) return createSignedOutUser();

  let response: Response;
  try {
    response = await apiFetch(path, {
      ...init,
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${token}`,
        ...init.headers,
      },
    });
  } catch (error) {
    throw error;
  }

  if (!response.ok) {
    throw new Error(await readErrorMessage(response));
  }

  const payload = (await response.json()) as AccountResponse;
  return mapServerProfile(payload.profile);
}

async function readErrorMessage(response: Response): Promise<string> {
  const contentType = response.headers.get('content-type') ?? '';
  if (contentType.includes('application/json')) {
    const payload = (await response.json()) as { error?: string };
    return friendlyErrorMessage(payload.error ?? response.statusText);
  }
  return friendlyErrorMessage(await response.text());
}
