import type { Session, User as SupabaseUser } from '@supabase/supabase-js';
import { fetchAccountProfile } from './accountApi';
import { createSignedOutUser, loadUser, saveUser, type UserProfile } from './auth';
export { isSupabaseConfigured, supabase } from './supabaseClient';
import { supabase } from './supabaseClient';

export function userFromSupabase(session: Session | null): UserProfile {
  if (!session?.user) return createSignedOutUser();
  return mergeSupabaseUser(loadUser(), session.user);
}

export function mergeSupabaseUser(current: UserProfile, supabaseUser: SupabaseUser): UserProfile {
  const email = supabaseUser.email ?? '';
  const provider = supabaseUser.app_metadata.provider === 'google' ? 'google' : 'email';
  const updated: UserProfile = {
    ...current,
    id: supabaseUser.id,
    email,
    displayName: supabaseUser.user_metadata.full_name ?? email.split('@')[0] ?? 'User',
    paymentStatus: 'active',
    authProvider: provider,
  };
  saveUser(updated);
  return updated;
}

export async function getCurrentAuthUser(): Promise<UserProfile> {
  if (!supabase) return loadUser();
  const { data, error } = await supabase.auth.getSession();
  if (error) throw new Error(error.message);
  if (!data.session?.user) return createSignedOutUser();
  return syncAccountProfile(data.session);
}

function validateEmailPassword(email: string, password: string): string {
  const normalizedEmail = email.trim().toLowerCase();
  if (!normalizedEmail || password.length < 6) throw new Error('Use an email and a password with at least 6 characters.');
  return normalizedEmail;
}

export async function signInWithEmail(email: string, password: string): Promise<UserProfile> {
  if (!supabase) throw new Error('Supabase is not configured. Add VITE_SUPABASE_URL and VITE_SUPABASE_ANON_KEY to .env.');
  const normalizedEmail = validateEmailPassword(email, password);

  const signIn = await supabase.auth.signInWithPassword({ email: normalizedEmail, password });
  if (signIn.error) throw new Error(signIn.error.message);
  if (signIn.data.session) return syncAccountProfile(signIn.data.session);
  if (signIn.data.user) return mergeSupabaseUser(loadUser(), signIn.data.user);
  throw new Error('Could not sign in. Check your email and password.');
}

export async function signUpWithEmail(email: string, password: string): Promise<{ user?: UserProfile; needsConfirmation: boolean }> {
  if (!supabase) throw new Error('Supabase is not configured. Add VITE_SUPABASE_URL and VITE_SUPABASE_ANON_KEY to .env.');
  const normalizedEmail = validateEmailPassword(email, password);

  const signUp = await supabase.auth.signUp({
    email: normalizedEmail,
    password,
    options: { emailRedirectTo: getAuthRedirectUrl() },
  });
  if (signUp.error) throw new Error(signUp.error.message);

  if (!signUp.data.user?.id) {
    throw new Error('Supabase did not create the account. Check that email signups are enabled and that the frontend uses the correct Supabase project.');
  }
  if (Array.isArray(signUp.data.user.identities) && signUp.data.user.identities.length === 0) {
    throw new Error('An account with this email may already exist. Use Sign in instead.');
  }

  if (signUp.data.session) {
    const user = await syncAccountProfile(signUp.data.session);
    return { user, needsConfirmation: false };
  }

  return { needsConfirmation: true };
}

export async function signInOrRegisterWithEmail(email: string, password: string): Promise<UserProfile> {
  const signUp = await signUpWithEmail(email, password);
  if (signUp.user) return signUp.user;
  throw new Error('Check your email to confirm the account, then sign in.');
}

export async function continueWithGoogle(): Promise<void> {
  if (!supabase) throw new Error('Supabase is not configured. Add VITE_SUPABASE_URL and VITE_SUPABASE_ANON_KEY to .env.');
  const { error } = await supabase.auth.signInWithOAuth({
    provider: 'google',
    options: { redirectTo: getAuthRedirectUrl() },
  });
  if (error) throw new Error(error.message);
}

function getAuthRedirectUrl(): string {
  return new URL(import.meta.env.BASE_URL, window.location.origin).toString();
}

export async function signOut(): Promise<UserProfile> {
  if (supabase) await supabase.auth.signOut();
  const signedOut = createSignedOutUser();
  saveUser(signedOut);
  return signedOut;
}

export function onAuthStateChanged(callback: (user: UserProfile) => void): () => void {
  if (!supabase) return () => undefined;
  const { data } = supabase.auth.onAuthStateChange((_event, session) => {
    if (!session?.user) {
      callback(createSignedOutUser());
      return;
    }
    window.setTimeout(() => {
      syncAccountProfile(session).then(callback);
    }, 0);
  });
  return () => data.subscription.unsubscribe();
}

async function syncAccountProfile(session: Session): Promise<UserProfile> {
  return fetchAccountProfile().catch(() => userFromSupabase(session));
}
