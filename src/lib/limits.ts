import type { DocumentAudit } from './captionUtils';
import { estimateCaptionCreditsNeeded, getPlan } from './access';
import type { UserProfile } from './auth';
import { getRemainingCaptionCredits, getRemainingDocuments, getRemainingExports } from './auth';

export interface LimitResult {
  allowed: boolean;
  message?: string;
}

export function canProcessDocument(user: UserProfile, assetCount: number): LimitResult {
  const plan = getPlan(user.planId);
  if (getRemainingDocuments(user) <= 0) {
    return { allowed: false, message: 'No document scans remaining this month.' };
  }
  if (assetCount > plan.maxAssetsPerDocument) {
    return { allowed: false, message: `This document has more than ${plan.maxAssetsPerDocument} figures/tables.` };
  }
  return { allowed: true };
}

export function canExportDocument(user: UserProfile): LimitResult {
  const plan = getPlan(user.planId);
  if (getRemainingExports(user) <= 0) {
    return { allowed: false, message: 'No exports remaining this month.' };
  }
  return { allowed: true };
}

export function canUseGeneratedCaptions(user: UserProfile, creditsNeeded: number): LimitResult {
  const plan = getPlan(user.planId);
  if (!plan.generatedCaptionsEnabled) {
    return { allowed: false, message: 'Generated captions are unavailable.' };
  }
  if (getRemainingCaptionCredits(user) < creditsNeeded) {
    return {
      allowed: false,
      message: getRemainingCaptionCredits(user) <= 0 ? 'You have no caption credits remaining.' : `This document needs ${creditsNeeded} caption credits; you have ${getRemainingCaptionCredits(user)}.`,
    };
  }
  return { allowed: true };
}

export function estimateAuditCaptionCredits(audit: DocumentAudit | null): number {
  if (!audit) return 0;
  return estimateCaptionCreditsNeeded(audit.totalFigures + audit.totalTables);
}
