// ============================================================================
// TrendChart   
// ============================================================================

import * as React from "react";
import { Box, Typography } from "@mui/material";
import { LineChart, Line, XAxis, YAxis, Tooltip, ResponsiveContainer } from "recharts";
import { TrendPoint } from "../types";

interface TrendChartProps {
  itemCode: string;
  points: TrendPoint[];
}

export default function TrendChart({ itemCode, points }: TrendChartProps) {
  if (!itemCode) {
    return (
      <Typography variant="caption" sx={{ fontSize: "11px", color: "#94a3b8" }}>
        Select an item to view its consumption trend.
      </Typography>
    );
  }

  if (points.length < 2) {
    return (
      <Typography variant="caption" sx={{ fontSize: "11px", color: "#94a3b8" }}>
        Not enough history yet for {itemCode} - trend builds up after a few sync cycles.
      </Typography>
    );
  }

  const chartData = points.map((p) => ({
    time: new Date(p.timestamp).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }),
    stock: p.stockLevel,
  }));

  return (
    <Box sx={{ width: "100%", height: 140 }}>
      <Typography variant="caption" sx={{ fontSize: "11px", fontWeight: 600, color: "#334155" }}>
        Stock Trend - {itemCode}
      </Typography>
      <ResponsiveContainer width="100%" height="100%">
        <LineChart data={chartData} margin={{ top: 8, right: 8, left: -20, bottom: 0 }}>
          <XAxis dataKey="time" tick={{ fontSize: 9 }} />
          <YAxis tick={{ fontSize: 9 }} />
          <Tooltip contentStyle={{ fontSize: "11px" }} />
          <Line type="monotone" dataKey="stock" stroke="#0062D6" strokeWidth={2} dot={false} />
        </LineChart>
      </ResponsiveContainer>
    </Box>
  );
}