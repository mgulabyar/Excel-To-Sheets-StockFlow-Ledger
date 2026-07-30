// ============================================================================
// auditLogger.ts - Full Traceability Ledger (Local/Excel Side)
// Fixes: only a single status cell was written, no real audit history
// ============================================================================

import { AuditEntry } from "../types";

const STORAGE_KEY = "stockflow_audit_log_v1";
const MAX_ENTRIES = 500; // keep the log bounded

function loadLog(): AuditEntry[] {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    return raw ? (JSON.parse(raw) as AuditEntry[]) : [];
  } catch {
    return [];
  }
}

function saveLog(log: AuditEntry[]): void {
  try {
    const trimmed = log.slice(-MAX_ENTRIES);
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(trimmed));
  } catch {
    // if storage fails, log still exists for current session via caller state
  }
}

export function recordAuditEntry(entry: Omit<AuditEntry, "id" | "timestamp">): AuditEntry {
  const fullEntry: AuditEntry = {
    ...entry,
    id: `audit-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    timestamp: new Date().toISOString(),
  };
  const log = loadLog();
  log.push(fullEntry);
  saveLog(log);
  return fullEntry;
}

export function getAuditLog(limit = 50): AuditEntry[] {
  const log = loadLog();
  return log.slice(-limit).reverse(); // most recent first
}

export function exportAuditLogAsCsv(): string {
  const log = loadLog();
  const header = "id,itemCode,action,beforeValue,afterValue,user,timestamp";
  const rows = log.map((e) =>
    [e.id, e.itemCode, e.action, e.beforeValue, e.afterValue, e.user, e.timestamp].join(",")
  );
  return [header, ...rows].join("\n");
}

export function clearAuditLog(): void {
  saveLog([]);
}