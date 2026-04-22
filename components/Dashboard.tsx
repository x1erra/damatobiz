'use client';

import { useState, useEffect, useMemo } from 'react';
import { fetchMetrics } from '@/lib/data';
import { aggregateByDate, calculateHealthScore } from '@/lib/analytics';
import { StatementMetric } from '@/types';
import UploadZone from './UploadZone';
import ProblemSpotter from './ProblemSpotter';
import DataTable from './DataTable';
import DataManagement from './DataManagement';
import dynamic from 'next/dynamic';
const Plot = dynamic(() => import('react-plotly.js'), { ssr: false });
import { HelpCircle } from 'lucide-react';
import clsx from 'clsx';

export default function Dashboard() {
    const [metrics, setMetrics] = useState<StatementMetric[]>([]);
    const [loading, setLoading] = useState(true);
    const [segmentFilter, setSegmentFilter] = useState<'all' | 'ip' | 'account'>('all');
    const [customDateRange, setCustomDateRange] = useState<{ start: string, end: string }>({ start: '', end: '' });

    // Calculate absolute data bounds
    const dataBounds = useMemo(() => {
        if (metrics.length === 0) return { min: '', max: '' };
        const dates = metrics.map(m => m.report_start_date).sort();
        const endDates = metrics.map(m => m.report_end_date).sort();
        return {
            min: dates[0],
            max: endDates[endDates.length - 1]
        };
    }, [metrics]);

    // Initialize date range to full bounds when data loads
    useEffect(() => {
        if (dataBounds.min && dataBounds.max && !customDateRange.start && !customDateRange.end) {
            setCustomDateRange({ start: dataBounds.min, end: dataBounds.max });
        }
    }, [dataBounds]);

    const refreshData = async () => {
        setLoading(true);
        try {
            const data = await fetchMetrics(99999);
            setMetrics(data);
        } catch (err) {
            console.error(err);
        } finally {
            setLoading(false);
        }
    };

    useEffect(() => {
        refreshData();
    }, []);


    const filteredMetrics = useMemo(() => {
        let result = [...metrics];
        if (customDateRange.start) result = result.filter(m => m.report_start_date >= customDateRange.start);
        if (customDateRange.end) result = result.filter(m => m.report_end_date <= customDateRange.end);

        if (segmentFilter === 'ip') {
            result = result.filter(m => m.statement_type.includes('IP-Level'));
        } else if (segmentFilter === 'account') {
            result = result.filter(m => m.statement_type.includes('Account Level'));
        }
        return result;
    }, [metrics, segmentFilter, customDateRange]);

    const aggregatedSeries = useMemo(() => aggregateByDate(filteredMetrics), [filteredMetrics]);

    const currentStats = useMemo(() => {
        if (filteredMetrics.length === 0) return { total: 0, health: 0, criticalPct: 0, avgLead: 0 };
        const total = filteredMetrics.reduce((acc, m) => acc + Number(m.request_count || 0), 0);
        const health = calculateHealthScore(filteredMetrics);
        const critical = filteredMetrics.filter(m => m.duration_bucket === "9. Over a minute").reduce((acc, m) => acc + Number(m.request_count || 0), 0);
        const leadTimeSum = filteredMetrics.reduce((acc, m) => acc + (Number(m.avg_lead_time || 0) * Number(m.request_count || 0)), 0);
        const avgLead = total > 0 ? leadTimeSum / total : 0;
        return { total, health, criticalPct: total > 0 ? (critical / total * 100) : 0, avgLead };
    }, [filteredMetrics]);

    const showDualBars = segmentFilter === 'all' &&
        metrics.some(m => m.statement_type.includes('IP-Level')) &&
        metrics.some(m => m.statement_type.includes('Account Level'));

    return (
        <div className="max-w-7xl mx-auto p-6 space-y-8">
            <header className="flex justify-between items-center mb-8">
                <div>
                    <h1 className="text-3xl font-bold bg-gradient-to-r from-sky-400 to-indigo-400 bg-clip-text text-transparent">
                        CI Stats Analyzer
                    </h1>
                    <p className="text-slate-400 font-medium">Performance Insights & Trend Analysis</p>
                </div>
            </header>

            <UploadZone onUploadComplete={refreshData} />

            {/* Controls */}
            <div className="flex flex-wrap gap-6 p-5 bg-slate-900/80 rounded-2xl border border-slate-800 backdrop-blur-md shadow-2xl">
                <div className="space-y-2">
                    <label className="text-[10px] text-slate-500 uppercase font-bold tracking-widest pl-1">Date Range Selection</label>
                    <div className="flex gap-2 items-center bg-slate-950/50 p-1.5 rounded-lg border border-slate-800">
                        <input type="date" value={customDateRange.start} min={dataBounds.min} max={customDateRange.end || dataBounds.max} onChange={(e) => setCustomDateRange(prev => ({ ...prev, start: e.target.value }))} className="bg-transparent text-slate-200 text-sm rounded px-2 py-1 focus:outline-none [color-scheme:dark]" />
                        <span className="text-slate-600 font-mono text-xs">→</span>
                        <input type="date" value={customDateRange.end} min={customDateRange.start || dataBounds.min} max={dataBounds.max} onChange={(e) => setCustomDateRange(prev => ({ ...prev, end: e.target.value }))} className="bg-transparent text-slate-200 text-sm rounded px-2 py-1 focus:outline-none [color-scheme:dark]" />
                        {(customDateRange.start !== dataBounds.min || customDateRange.end !== dataBounds.max) && (
                            <button onClick={() => setCustomDateRange({ start: dataBounds.min, end: dataBounds.max })} className="text-[10px] text-sky-400 hover:text-sky-300 ml-2 font-bold uppercase tracking-tight pr-2">Reset</button>
                        )}
                    </div>
                </div>
                <div className="space-y-2">
                    <label className="text-[10px] text-slate-500 uppercase font-bold tracking-widest pl-1">Statement Segment</label>
                    <div className="flex rounded-lg bg-slate-950/50 p-1 border border-slate-800">
                        {[{ id: 'all', label: 'All Segments' }, { id: 'ip', label: 'IP Level' }, { id: 'account', label: 'Account Level' }].map((s) => (
                            <button key={s.id} onClick={() => setSegmentFilter(s.id as any)} className={clsx("px-4 py-1.5 text-xs font-semibold rounded-md transition-all duration-200", segmentFilter === s.id ? "bg-sky-500 text-white shadow-lg shadow-sky-500/20" : "text-slate-500 hover:text-slate-300")}>{s.label}</button>
                        ))}
                    </div>
                </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
                <KPICard label="Total Requests" value={currentStats.total.toLocaleString()} subtext="Captured Volume" tooltip="Sum of all generate requests within the current filters." />
                <KPICard
                    label="Health Score"
                    value={`${currentStats.health.toFixed(1)}%`}
                    trend={currentStats.health >= 70 ? 'up' : 'down'}
                    subtext="Requests under 20 sec"
                    tooltip="Percent of requests completing in <20s."
                    customColor={getHealthColor(currentStats.health)}
                />
                <KPICard
                    label="Critical Wait"
                    value={`${currentStats.criticalPct.toFixed(1)}%`}
                    trend={currentStats.criticalPct < 15 ? 'down' : 'up'}
                    isInverse
                    subtext="Requests over 1 min"
                    tooltip="Percent of requests taking >1m. Lower is better."
                    customColor={getCriticalColor(currentStats.criticalPct)}
                />
                <KPICard label="Avg Lead Time" value={`${currentStats.avgLead.toFixed(1)}s`} subtext="Weighted Avg" tooltip="Average generation time across all selected segments." />
            </div>

            <ProblemSpotter metrics={filteredMetrics} />

            {/* Charts Section */}
            {aggregatedSeries.length > 0 && (
                <div className="space-y-6">
                    {/* Line Chart - Trend (Full Width if dual) */}
                    <div className={clsx("bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden h-[450px] shadow-xl", showDualBars ? "w-full" : "w-full lg:w-[calc(50%-12px)] inline-block align-top")}>
                        <div className="p-6 pb-2"><h3 className="text-sm font-bold uppercase tracking-widest text-slate-400 flex items-center gap-2"><div className="w-2 h-2 rounded-full bg-emerald-500" />Health Score Trend</h3></div>
                        <Plot
                            data={segmentFilter === 'all' ? [
                                { x: aggregatedSeries.map(d => d.date), y: aggregatedSeries.map(d => calculateHealthScore(d.metrics.filter(m => m.statement_type.includes('IP-Level')))), type: 'scatter', mode: 'lines+markers', name: 'IP Level', line: { color: '#0ea5e9', width: 2, shape: 'spline' }, marker: { size: 4 }, hovertemplate: 'IP: %{y:.1f}%<extra></extra>' },
                                { x: aggregatedSeries.map(d => d.date), y: aggregatedSeries.map(d => calculateHealthScore(d.metrics.filter(m => m.statement_type.includes('Account Level')))), type: 'scatter', mode: 'lines+markers', name: 'Account Level', line: { color: '#8b5cf6', width: 2, shape: 'spline' }, marker: { size: 4 }, hovertemplate: 'Acc: %{y:.1f}%<extra></extra>' }
                            ] : [
                                { x: aggregatedSeries.map(d => d.date), y: aggregatedSeries.map(d => d.healthScore), type: 'scatter', mode: 'lines+markers', line: { color: '#10b981', width: 2, shape: 'spline' }, marker: { size: 6, color: '#10b981' }, hovertemplate: 'Health: %{y:.1f}%<extra></extra>' }
                            ]}
                            layout={{
                                autosize: true, paper_bgcolor: 'transparent', plot_bgcolor: 'transparent', margin: { t: 20, r: 30, l: 50, b: 80 },
                                xaxis: { gridcolor: '#1e293b', tickfont: { color: '#64748b', size: 9 }, tickangle: -45, type: 'date', tickformat: '%b %d', automargin: true },
                                yaxis: { gridcolor: '#1e293b', tickfont: { color: '#64748b', size: 9 }, range: [0, 105], showgrid: true, fixedrange: true },
                                showlegend: segmentFilter === 'all', legend: { font: { color: '#94a3b8', size: 10 }, orientation: 'h', y: -0.3, x: 0.5, xanchor: 'center' }
                            }}
                            config={{ responsive: true, displayModeBar: false }}
                            style={{ width: '100%', height: 'calc(100% - 60px)' }}
                        />
                    </div>

                    {!showDualBars && (
                        <div className="w-full lg:w-[calc(50%-12px)] inline-block align-top ml-6">
                            <PerformanceBarCard title="Performance Volume" metricsSeries={filteredMetrics} dates={aggregatedSeries.map(d => d.date)} />
                        </div>
                    )}

                    {showDualBars && (
                        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                            <PerformanceBarCard title="IP Level Distribution" metricsSeries={filteredMetrics.filter(m => m.statement_type.includes('IP-Level'))} dates={aggregatedSeries.map(d => d.date)} />
                            <PerformanceBarCard title="Account Level Distribution" metricsSeries={filteredMetrics.filter(m => m.statement_type.includes('Account Level'))} dates={aggregatedSeries.map(d => d.date)} />
                        </div>
                    )}
                </div>
            )}

            <DataTable metrics={filteredMetrics} />
            <DataManagement onDataChange={refreshData} />
        </div>
    );
}

function PerformanceBarCard({ title, metricsSeries, dates }: any) {
    return (
        <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden h-[400px] shadow-xl">
            <div className="p-6 pb-2"><h3 className="text-sm font-bold uppercase tracking-widest text-slate-400 flex items-center gap-2"><div className="w-2 h-2 rounded-full bg-sky-500" />{title}</h3></div>
            <Plot
                data={[
                    { x: dates, y: dates.map((date: string) => metricsSeries.filter((m: any) => m.report_start_date === date && ["1. less than 5 secs", "2. less than 10 secs", "3. less than 15 secs", "4. less than 20 secs", "5. less than 25 secs", "6. less than 30 secs"].includes(m.duration_bucket)).reduce((a: any, b: any) => a + Number(b.request_count || 0), 0)), name: 'Fast', type: 'bar', marker: { color: '#10b981' } },
                    { x: dates, y: dates.map((date: string) => metricsSeries.filter((m: any) => m.report_start_date === date && ["7. less than 45 secs", "8. less than 1 min"].includes(m.duration_bucket)).reduce((a: any, b: any) => a + Number(b.request_count || 0), 0)), name: 'Moderate', type: 'bar', marker: { color: '#f59e0b' } },
                    { x: dates, y: dates.map((date: string) => metricsSeries.filter((m: any) => m.report_start_date === date && m.duration_bucket === "9. Over a minute").reduce((a: any, b: any) => a + Number(b.request_count || 0), 0)), name: 'Critical', type: 'bar', marker: { color: '#ef4444' } }
                ]}
                layout={{
                    barmode: 'stack', autosize: true, paper_bgcolor: 'transparent', plot_bgcolor: 'transparent', margin: { t: 20, r: 30, l: 50, b: 80 },
                    xaxis: { gridcolor: '#1e293b', tickfont: { color: '#64748b', size: 9 }, tickangle: -45, type: 'date', tickformat: '%b %d', automargin: true },
                    yaxis: { gridcolor: '#1e293b', tickfont: { color: '#64748b', size: 9 }, showgrid: true, fixedrange: true },
                    legend: { font: { color: '#94a3b8', size: 9 }, orientation: 'h', y: -0.3, x: 0.5, xanchor: 'center' }
                }}
                config={{ responsive: true, displayModeBar: false }}
                style={{ width: '100%', height: 'calc(100% - 60px)' }}
            />
        </div>
    );
}

function Tooltip({ text }: { text: string }) {
    if (!text) return null;
    return (
        <div className="absolute bottom-full left-1/2 -translate-x-1/2 mb-3 p-2.5 bg-slate-800 text-slate-200 text-[11px] rounded-lg shadow-2xl border border-slate-700 w-52 invisible group-hover:visible z-[100] transition-all opacity-0 group-hover:opacity-100 backdrop-blur-md pointer-events-none leading-relaxed">
            {text}<div className="absolute top-full left-1/2 -translate-x-1/2 border-[6px] border-transparent border-t-slate-800" />
        </div>
    );
}

function getHealthColor(score: number): string {
    if (score < 50) {
        // Red Range: 0 (worst/brightest) -> 49 (least bad/dullest)
        const intensity = (49 - score) / 49;
        const lightness = 35 + (intensity * 30); // 35% (dull red) to 65% (bright red)
        return `hsl(0, 85%, ${lightness}%)`;
    } else if (score < 70) {
        // Yellow Range: 50 (worst/brightest) -> 69 (best/dullest)
        const intensity = (69 - score) / (69 - 50);
        const lightness = 35 + (intensity * 25); // 35% (dull yellow) to 60% (bright yellow)
        return `hsl(45, 95%, ${lightness}%)`;
    } else {
        // Green Range: 70 (worst/dullest) -> 100 (best/brightest)
        const intensity = (score - 70) / (100 - 70);
        const lightness = 35 + (intensity * 30); // 35% (dull green) to 65% (bright green)
        return `hsl(150, 80%, ${lightness}%)`;
    }
}

function getCriticalColor(score: number): string {
    if (score < 10) {
        // Green Range: 0 (best/brightest) -> 10 (dullest)
        const intensity = (10 - score) / 10;
        const lightness = 35 + (intensity * 30);
        return `hsl(150, 80%, ${lightness}%)`;
    } else if (score < 25) {
        // Yellow Range: 10 (dullest) -> 25 (worst/brightest)
        const intensity = (score - 10) / (25 - 10);
        const lightness = 35 + (intensity * 25);
        return `hsl(45, 95%, ${lightness}%)`;
    } else {
        // Red Range: 25 (dullest) -> 100 (worst/brightest)
        const intensity = (score - 25) / 75;
        const lightness = 35 + (intensity * 30);
        return `hsl(0, 85%, ${lightness}%)`;
    }
}

function KPICard({ label, value, trend, isInverse, subtext, tooltip, customColor }: any) {
    let colorClass = "text-slate-100";
    if (trend && !customColor) {
        if (isInverse) colorClass = trend === 'up' ? "text-rose-400" : "text-emerald-400";
        else colorClass = trend === 'up' ? "text-emerald-400" : "text-rose-400";
    }
    return (
        <div className="bg-slate-900/40 border border-slate-800/60 p-6 rounded-2xl relative group hover:bg-slate-900/60 hover:border-slate-700 transition-all cursor-default">
            <h4 className="text-slate-500 text-[10px] font-bold uppercase tracking-widest mb-2 flex items-center gap-1.5">{label}<HelpCircle className="w-3 h-3 text-slate-700" /></h4>
            <div
                className={clsx("text-3xl font-bold tracking-tight", !customColor && colorClass)}
                style={customColor ? { color: customColor } : {}}
            >
                {value}
            </div>
            <p className="text-[10px] text-slate-600 mt-2 font-bold uppercase tracking-widest">{subtext}</p>
            <Tooltip text={tooltip} />
        </div>
    )
}
