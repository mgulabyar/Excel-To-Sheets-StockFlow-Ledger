// ============================================================================
// syncQueue.ts - Resilient Offline Retry Queue
// Fixes: failed pushes were silently lost (only console.error, no retry)
// ============================================================================

import { QueuedRequest, WritebackPayload } from "../types";
import { sendSecurePost } from "./secureRequest";

const STORAGE_KEY = "stockflow_sync_queue_v1";
const MAX_ATTEMPTS = 5;

function loadQueue(): QueuedRequest[] {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    return raw ? (JSON.parse(raw) as QueuedRequest[]) : [];
  } catch {
    return [];
  }
}

function saveQueue(queue: QueuedRequest[]): void {
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(queue));
  } catch {
    // storage unavailable - queue will only live in memory for this session
  }
}

/**
 * Adds a failed writeback payload to the retry queue, in original order.
 */
export function enqueueFailedRequest(payload: WritebackPayload): QueuedRequest {
  const queue = loadQueue();
  const entry: QueuedRequest = {
    id: `${payload.itemCode}-${Date.now()}`,
    payload,
    createdAt: new Date().toISOString(),
    attempts: 0,
  };
  queue.push(entry);
  saveQueue(queue);
  return entry;
}

export function getQueueLength(): number {
  return loadQueue().length;
}

export function getQueueSnapshot(): QueuedRequest[] {
  return loadQueue();
}

/**
 * Attempts to flush the entire queue in FIFO order. Stops retrying an
 * individual item after MAX_ATTEMPTS to avoid infinite retry storms, and
 * reports how many succeeded / how many remain.
 */
export async function flushQueue(
  apiUrl: string
): Promise<{ succeeded: number; remaining: number; failed: number }> {
  let queue = loadQueue();
  if (queue.length === 0) {
    return { succeeded: 0, remaining: 0, failed: 0 };
  }

  let succeeded = 0;
  let failed = 0;
  const stillPending: QueuedRequest[] = [];

  for (const entry of queue) {
    const result = await sendSecurePost(apiUrl, entry.payload as unknown as Record<string, any>);

    if (result.ok) {
      succeeded += 1;
      continue;
    }

    entry.attempts += 1;
    if (entry.attempts >= MAX_ATTEMPTS) {
      failed += 1; // permanently dropped after too many failures
    } else {
      stillPending.push(entry);
    }
  }

  saveQueue(stillPending);
  return { succeeded, remaining: stillPending.length, failed };
}

export function clearQueue(): void {
  saveQueue([]);
}