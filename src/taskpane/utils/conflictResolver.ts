import { ConflictRecord } from "../types";

export type ConflictStrategy = "last_write_wins" | "manual_for_flagged";


export function detectConflict(params: {
  itemCode: string;
  localValue: number;
  localTimestamp: string;
  cloudValue: number;
  cloudTimestamp: string;
  lastSyncedTimestamp: string | null;
}): ConflictRecord | null {
  const { itemCode, localValue, localTimestamp, cloudValue, cloudTimestamp, lastSyncedTimestamp } = params;

  if (localValue === cloudValue) return null;

  const localEditedAfterSync =
    !lastSyncedTimestamp || new Date(localTimestamp) > new Date(lastSyncedTimestamp);
  const cloudEditedAfterSync =
    !lastSyncedTimestamp || new Date(cloudTimestamp) > new Date(lastSyncedTimestamp);

  if (!(localEditedAfterSync && cloudEditedAfterSync)) return null;

  return {
    itemCode,
    localValue,
    localTimestamp,
    cloudValue,
    cloudTimestamp,
    resolution: null,
  };
}


export function resolveConflict(
  conflict: ConflictRecord,
  strategy: ConflictStrategy,
  isFlaggedItem: boolean
): ConflictRecord {
  if (strategy === "manual_for_flagged" && isFlaggedItem) {
    return { ...conflict, resolution: "manual_pending" };
  }

  const localIsNewer = new Date(conflict.localTimestamp) > new Date(conflict.cloudTimestamp);
  return { ...conflict, resolution: localIsNewer ? "local_wins" : "cloud_wins" };
}