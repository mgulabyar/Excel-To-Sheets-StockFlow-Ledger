import { TrendPoint, VendorLeadTimeRecord } from "../types";

const TREND_KEY = "stockflow_trend_points_v1";
const LEADTIME_KEY = "stockflow_leadtime_records_v1";
const MAX_TREND_POINTS_PER_ITEM = 60; 

function loadJson<T>(key: string): T[] {
  try {
    const raw = window.localStorage.getItem(key);
    return raw ? (JSON.parse(raw) as T[]) : [];
  } catch {
    return [];
  }
}

function saveJson<T>(key: string, data: T[]): void {
  try {
    window.localStorage.setItem(key, JSON.stringify(data));
  } catch {
  }
}

/**
 * Records a stock-level observation for an item, used to build the
 * consumption trend chart.
 */
export function recordTrendPoint(itemCode: string, stockLevel: number): void {
  const points = loadJson<TrendPoint>(TREND_KEY);
  points.push({ itemCode, stockLevel, timestamp: new Date().toISOString() });

  // keep only the most recent N points per item to bound storage size
  const forItem = points.filter((p) => p.itemCode === itemCode);
  if (forItem.length > MAX_TREND_POINTS_PER_ITEM) {
    const overflow = forItem.length - MAX_TREND_POINTS_PER_ITEM;
    let removed = 0;
    const filtered = points.filter((p) => {
      if (p.itemCode === itemCode && removed < overflow) {
        removed += 1;
        return false;
      }
      return true;
    });
    saveJson(TREND_KEY, filtered);
    return;
  }
  saveJson(TREND_KEY, points);
}

export function getTrendPoints(itemCode: string): TrendPoint[] {
  return loadJson<TrendPoint>(TREND_KEY).filter((p) => p.itemCode === itemCode);
}

/**
 * Records how long a vendor actually took to deliver vs. what was promised,
 * for a given item. This feeds the adaptive lead-time formula below.
 */
export function recordVendorDelivery(
  itemCode: string,
  promisedLeadTime: number,
  actualLeadTime: number
): void {
  const records = loadJson<VendorLeadTimeRecord>(LEADTIME_KEY);
  records.push({ itemCode, promisedLeadTime, actualLeadTime, recordedAt: new Date().toISOString() });
  saveJson(LEADTIME_KEY, records);
}

/**
 * Adaptive Lead-Time Learning:
 * Instead of trusting the static "promised" lead time in the stockout
 * formula, this blends it with the historical actual average, so the
 * reorder threshold gets more accurate over time without manual tuning.
 *
 * Uses a simple weighted average (70% historical actual, 30% promised)
 * once at least 3 data points exist; otherwise falls back to promised.
 */
export function getAdaptiveLeadTime(itemCode: string, promisedLeadTime: number): number {
  const records = loadJson<VendorLeadTimeRecord>(LEADTIME_KEY).filter((r) => r.itemCode === itemCode);
  if (records.length < 3) return promisedLeadTime;

  const avgActual =
    records.reduce((sum, r) => sum + r.actualLeadTime, 0) / records.length;

  return Number((avgActual * 0.7 + promisedLeadTime * 0.3).toFixed(2));
}