export type PlanId = 'free';

export interface AccessPlan {
  id: PlanId;
  name: string;
  audience: string;
  documentLimit: number;
  maxAssetsPerDocument: number;
  captionCredits: number;
  teamSeats: number;
  exportLimit: number;
  generatedCaptionsEnabled: boolean;
  documentAudit: boolean;
  captionImprovement: boolean;
  batchProcessing: boolean;
  features: string[];
}

export const ACCESS_PLANS: Record<PlanId, AccessPlan> = {
  free: {
    id: 'free',
    name: 'Free',
    audience: 'Complete document automation at no cost',
    documentLimit: 1000,
    maxAssetsPerDocument: 200,
    captionCredits: 5000,
    teamSeats: 1,
    exportLimit: 1000,
    generatedCaptionsEnabled: true,
    documentAudit: true,
    captionImprovement: true,
    batchProcessing: true,
    features: [
      'All document automation features',
      'Generated captions',
      'Manual captions always available',
      'Caption improvement',
      'Automatic figure and table detection',
      'Automated Table of Contents',
      'Automated Table of Figures',
      'Automated Table of Tables',
      'Export corrected Word documents',
    ],
  },
};

export function getPlan(planId: PlanId): AccessPlan {
  return ACCESS_PLANS[planId];
}

export function normalizePlanId(_planId: string | undefined): PlanId {
  return 'free';
}

export function estimateCaptionCreditsNeeded(totalAssets: number): number {
  return Math.max(0, totalAssets);
}
