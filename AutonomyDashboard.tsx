"use client";
import React, { useState, useEffect } from "react";
import { Card, CardHeader, CardTitle, CardContent } from "./ui/card";
import { mockData } from "../lib/mock-data";
import {
  ResponsiveContainer,
  LineChart,
  Line,
  PieChart,
  Pie,
  Cell,
  XAxis,
  YAxis,
  Tooltip,
  Legend,
  BarChart,
  Bar,
} from "recharts";

const COLORS = ["#28a745", "#ffc107", "#dc3545", "#17a2b8"];

function formatUTCDateTime(date: Date) {
  return date.toISOString().replace("T", " ").slice(0, 19);
}

export default function AutonomyDashboard() {
  const [currentTime, setCurrentTime] = useState(formatUTCDateTime(new Date()));
  const [showDetailedHealth, setShowDetailedHealth] = useState(false);

  // Use mockData for this demo
  const data = mockData;
  const healthScore = Math.round(
    data.healthBreakdown.reduce((acc, val) => acc + val.value, 0) /
      data.healthBreakdown.length
  );

  useEffect(() => {
    const timer = setInterval(() => {
      setCurrentTime(formatUTCDateTime(new Date()));
    }, 1000);
    return () => clearInterval(timer);
  }, []);

  return (
    <div className="min-h-screen bg-gray-100 p-6">
      {/* Header */}
      <div className="flex justify-between items-center bg-white rounded-lg shadow p-4 mb-6">
        <div>
          <h1 className="text-2xl font-bold mb-1">Autonomy Health Dashboard</h1>
          <div className="text-gray-600">
            Current Date and Time (UTC): {currentTime}
          </div>
        </div>
        <div className="text-right">
          <div className="font-medium">User: JGipson9</div>
          <div className="text-sm text-gray-500">Connected to Production System</div>
        </div>
      </div>

      {/* Top Row: Health Score & Trend */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
        {/* Overall Health Circle */}
        <Card
          className="cursor-pointer"
          onClick={() => setShowDetailedHealth((v) => !v)}
        >
          <CardHeader>
            <CardTitle>Overall Health Score</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="relative h-32 w-32 mx-auto flex items-center justify-center">
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie
                    data={[
                      { value: healthScore },
                      { value: 100 - healthScore },
                    ]}
                    dataKey="value"
                    innerRadius={50}
                    outerRadius={60}
                    startAngle={90}
                    endAngle={-270}
                  >
                    <Cell fill={healthScore >= 95 ? COLORS[0] : COLORS[2]} />
                    <Cell fill="#e5e7eb" />
                  </Pie>
                </PieChart>
              </ResponsiveContainer>
              <div className="absolute text-3xl font-bold">
                {healthScore}%
              </div>
            </div>
            <div className="text-center text-sm text-gray-500 mt-2">
              Click for breakdown
            </div>
          </CardContent>
        </Card>

        {/* 7-Day Health Trend */}
        <Card className="md:col-span-2">
          <CardHeader>
            <CardTitle>Health Trend (7 Days)</CardTitle>
          </CardHeader>
          <CardContent>
            <ResponsiveContainer width="100%" height={180}>
              <LineChart data={data.healthTrend}>
                <XAxis dataKey="date" />
                <YAxis domain={[0, 100]} />
                <Tooltip />
                <Line type="monotone" dataKey="health" stroke={COLORS[0]} />
              </LineChart>
            </ResponsiveContainer>
          </CardContent>
        </Card>
      </div>

      {/* Detailed Health Breakdown (Expandable) */}
      {showDetailedHealth && (
        <Card className="mb-4">
          <CardHeader>
            <CardTitle>Health Score Breakdown</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <ResponsiveContainer width="100%" height={220}>
                <PieChart>
                  <Pie
                    dataKey="value"
                    data={data.healthBreakdown}
                    label
                    outerRadius={80}
                  >
                    {data.healthBreakdown.map((entry, idx) => (
                      <Cell key={entry.name} fill={COLORS[idx % COLORS.length]} />
                    ))}
                  </Pie>
                  <Tooltip />
                  <Legend />
                </PieChart>
              </ResponsiveContainer>
              <div className="flex flex-col justify-center">
                {data.healthBreakdown.map((item) => (
                  <div
                    key={item.name}
                    className="flex justify-between items-center mb-2"
                  >
                    <span className="font-medium">{item.name}</span>
                    <span>{item.value}%</span>
                  </div>
                ))}
              </div>
            </div>
          </CardContent>
        </Card>
      )}

      {/* Vehicle Performance */}
      <Card>
        <CardHeader>
          <CardTitle>Moves per Hour by Vehicle</CardTitle>
        </CardHeader>
        <CardContent>
          <ResponsiveContainer width="100%" height={180}>
            <BarChart data={data.movesPerHour}>
              <XAxis dataKey="vehicle" />
              <YAxis />
              <Tooltip />
              <Legend />
              <Bar dataKey="moves" fill={COLORS[3]} />
              <Bar dataKey="average" fill={COLORS[1]} />
            </BarChart>
          </ResponsiveContainer>
        </CardContent>
      </Card>
    </div>
  );
}