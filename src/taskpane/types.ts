// ============================================================================
// StockFlow Ledger - Shared Types (Excel Side)
// ============================================================================

export interface CloudInventoryPayload {
  itemCode: string;
  physicalStock: number;
  expectedCargo: number;
  leadTime: number;
  dailySalesRate: number;
  lastUpdated?: string; // ISO timestamp - needed for conflict resolution
  updatedBy?: string;
}

export interface WritebackPayload {
  itemCode: string;
  newQuantity: number;
  timestamp: string;
  triggeredBy: string;
  approvalStage: ApprovalStage;
}

export type ApprovalStage = "draft" | "pending_finance" | "approved" | "rejected";

export type SyncStatus = "idle" | "syncing" | "synced" | "conflict" | "offline" | "error";

export interface QueuedRequest {
  id: string;
  payload: WritebackPayload;
  createdAt: string;
  attempts: number;
}

export interface ConflictRecord {
  itemCode: string;
  localValue: number;
  localTimestamp: string;
  cloudValue: number;
  cloudTimestamp: string;
  resolution: "local_wins" | "cloud_wins" | "manual_pending" | null;
}

export interface AuditEntry {
  id: string;
  itemCode: string;
  action: "push_approval" | "sync_pull" | "conflict_resolved" | "approval_stage_change";
  beforeValue: string | number | null;
  afterValue: string | number | null;
  user: string;
  timestamp: string;
}

export interface TrendPoint {
  itemCode: string;
  timestamp: string;
  stockLevel: number;
}

export interface VendorLeadTimeRecord {
  itemCode: string;
  promisedLeadTime: number;
  actualLeadTime: number;
  recordedAt: string;
}