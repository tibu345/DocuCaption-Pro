import { getPlan, normalizePlanId, type PlanId } from './access';

export interface UserUsage {
  documentsProcessed: number;
  exportsUsed: number;
  captionCreditsRemaining: number;
  failedProcessingJobs: number;
  periodStartedAt: string;
  billingPeriodEnd: string;
}

export interface UserProfile {
  id: string;
  email: string;
  displayName: string;
  planId: PlanId;
  paymentStatus: 'signed_out' | 'inactive' | 'active' | 'past_due' | 'canceled';
  usage: UserUsage;
  isAdmin: boolean;
  authProvider: 'none' | 'email' | 'google';
}

const STORAGE_KEY = 'docucaption:user';

export function createSignedOutUser(): UserProfile {
  return {
    id: 'signed-out',
    email: '',
    displayName: 'Signed out',
    planId: 'free',
    paymentStatus: 'signed_out',
    usage: emptyUsage(),
    isAdmin: false,
    authProvider: 'none',
  };
}

export function loadUser(): UserProfile {
  const stored = localStorage.getItem(STORAGE_KEY);
  if (!stored) {
    const signedOut = createSignedOutUser();
    saveUser(signedOut);
    return signedOut;
  }

  try {
    return resetMonthlyUsageIfNeeded(normalizeStoredUser(JSON.parse(stored) as Partial<UserProfile>));
  } catch {
    const signedOut = createSignedOutUser();
    saveUser(signedOut);
    return signedOut;
  }
}

export function saveUser(user: UserProfile): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(user));
}

export function signInWithEmail(email: string, password: string): UserProfile {
  const normalizedEmail = email.trim().toLowerCase();
  if (!normalizedEmail || password.length < 6) {
    throw new Error('Use an email and a password with at least 6 characters.');
  }

  const user: UserProfile = {
    ...loadUser(),
    id: `user-${btoa(normalizedEmail).replace(/=+$/g, '')}`,
    email: normalizedEmail,
    displayName: normalizedEmail.split('@')[0],
    paymentStatus: 'active',
    authProvider: 'email',
  };
  saveUser(user);
  return user;
}

export function signInWithGoogle(): UserProfile {
  const user: UserProfile = {
    ...loadUser(),
    id: 'google-local-user',
    email: 'google.user@docucaption.local',
    displayName: 'Google User',
    paymentStatus: 'active',
    authProvider: 'google',
  };
  saveUser(user);
  return user;
}

export function isAuthenticated(user: UserProfile): boolean {
  return user.authProvider !== 'none' && Boolean(user.email);
}

export function resetCredits(user: UserProfile): UserProfile {
  const updated = { ...user, usage: emptyUsage() };
  saveUser(updated);
  return updated;
}

export function setLocalTestingPlan(user: UserProfile, planId: PlanId): UserProfile {
  const plan = getPlan(planId);
  const now = new Date();
  const periodEnd = new Date(now);
  periodEnd.setUTCMonth(periodEnd.getUTCMonth() + 1);
  const updated: UserProfile = {
    ...user,
    planId,
    usage: {
      ...user.usage,
      documentsProcessed: 0,
      exportsUsed: 0,
      captionCreditsRemaining: plan.captionCredits,
      periodStartedAt: now.toISOString(),
      billingPeriodEnd: periodEnd.toISOString(),
    },
  };
  saveUser(updated);
  return updated;
}

export function trackProcessingFailure(user: UserProfile): UserProfile {
  return persistUsage(user, { failedProcessingJobs: user.usage.failedProcessingJobs + 1 });
}

export function consumeLocalDocument(user: UserProfile): UserProfile {
  return persistUsage(user, { documentsProcessed: user.usage.documentsProcessed + 1 });
}

export function consumeLocalExport(user: UserProfile): UserProfile {
  return persistUsage(user, { exportsUsed: user.usage.exportsUsed + 1 });
}

export function consumeLocalCaptionCredits(user: UserProfile, credits: number): UserProfile {
  return persistUsage(user, { captionCreditsRemaining: Math.max(0, user.usage.captionCreditsRemaining - credits) });
}

export function getRemainingDocuments(user: UserProfile): number {
  return Math.max(0, getPlan(user.planId).documentLimit - user.usage.documentsProcessed);
}

export function getRemainingExports(user: UserProfile): number {
  return Math.max(0, getPlan(user.planId).exportLimit - user.usage.exportsUsed);
}

export function getRemainingCaptionCredits(user: UserProfile): number {
  return Math.max(0, user.usage.captionCreditsRemaining);
}

function persistUsage(user: UserProfile, usage: Partial<UserUsage>): UserProfile {
  const updated = { ...user, usage: { ...user.usage, ...usage } };
  saveUser(updated);
  return updated;
}

function resetMonthlyUsageIfNeeded(user: UserProfile): UserProfile {
  const periodStarted = new Date(user.usage.periodStartedAt);
  const now = new Date();
  if (periodStarted.getUTCFullYear() === now.getUTCFullYear() && periodStarted.getUTCMonth() === now.getUTCMonth()) {
    return user;
  }
  const updated = { ...user, usage: emptyUsage() };
  saveUser(updated);
  return updated;
}

function normalizeStoredUser(user: Partial<UserProfile>): UserProfile {
  const fallback = createSignedOutUser();
  const legacyPlanId: string = typeof user.planId === 'string' ? user.planId : 'free';
  const planId = normalizePlanId(legacyPlanId);
  const legacyStatus = user.paymentStatus as UserProfile['paymentStatus'] | 'guest' | undefined;
  const authProvider = user.authProvider ?? (user.email ? 'email' : 'none');
  const storedUsage = user.usage as Partial<UserUsage> & { aiCreditsUsed?: number };
  return {
    ...fallback,
    ...user,
    id: authProvider === 'none' ? fallback.id : user.id ?? fallback.id,
    email: authProvider === 'none' ? '' : user.email ?? '',
    displayName: authProvider === 'none' ? fallback.displayName : user.displayName ?? user.email?.split('@')[0] ?? fallback.displayName,
    planId,
    paymentStatus: authProvider === 'none' ? 'signed_out' : legacyStatus === 'guest' ? 'active' : legacyStatus ?? 'active',
    usage: {
      ...emptyUsage(),
      ...storedUsage,
      captionCreditsRemaining:
        typeof storedUsage.captionCreditsRemaining === 'number'
          ? storedUsage.captionCreditsRemaining
          : Math.max(0, getPlan(planId).captionCredits - (storedUsage.aiCreditsUsed ?? 0)),
    },
    authProvider,
    isAdmin: Boolean(user.isAdmin),
  };
}

function emptyUsage(): UserUsage {
  const now = new Date();
  const periodEnd = new Date(now);
  periodEnd.setUTCMonth(periodEnd.getUTCMonth() + 1);
  return {
    documentsProcessed: 0,
    exportsUsed: 0,
    captionCreditsRemaining: getPlan('free').captionCredits,
    failedProcessingJobs: 0,
    periodStartedAt: now.toISOString(),
    billingPeriodEnd: periodEnd.toISOString(),
  };
}
