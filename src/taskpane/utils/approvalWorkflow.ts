// ============================================================================
// approvalWorkflow.ts - Two-Step Approval Chain
// Fixes: a single click was going straight to "approved", no second check
// ============================================================================

import { ApprovalStage } from "../types";

interface TransitionResult {
  nextStage: ApprovalStage;
  allowed: boolean;
  message: string;
}

/**
 * Defines valid transitions between approval stages.
 * draft -> pending_finance (Procurement Director submits)
 * pending_finance -> approved (Finance Manager confirms)
 * pending_finance -> rejected (Finance Manager rejects)
 * approved/rejected are terminal states for a given writeback.
 */
export function submitForFinanceApproval(currentStage: ApprovalStage): TransitionResult {
  if (currentStage !== "draft") {
    return {
      nextStage: currentStage,
      allowed: false,
      message: "Only draft items can be submitted for finance approval.",
    };
  }
  return {
    nextStage: "pending_finance",
    allowed: true,
    message: "Submitted. Awaiting Finance Manager approval.",
  };
}

export function financeApprove(currentStage: ApprovalStage): TransitionResult {
  if (currentStage !== "pending_finance") {
    return {
      nextStage: currentStage,
      allowed: false,
      message: "Only items pending finance review can be approved.",
    };
  }
  return {
    nextStage: "approved",
    allowed: true,
    message: "Approved. Writeback authorized for cloud sync.",
  };
}

export function financeReject(currentStage: ApprovalStage): TransitionResult {
  if (currentStage !== "pending_finance") {
    return {
      nextStage: currentStage,
      allowed: false,
      message: "Only items pending finance review can be rejected.",
    };
  }
  return {
    nextStage: "rejected",
    allowed: true,
    message: "Rejected. Writeback blocked.",
  };
}

export function isReadyToSync(stage: ApprovalStage): boolean {
  return stage === "approved";
}

export function stageLabel(stage: ApprovalStage): string {
  const labels: Record<ApprovalStage, string> = {
    draft: "Draft",
    pending_finance: "Awaiting Finance Approval",
    approved: "Approved - Synced",
    rejected: "Rejected",
  };
  return labels[stage];
}