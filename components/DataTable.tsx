'use client';

import { StatementMetric } from '@/types';
import { useState, useMemo, useEffect } from 'react';
import clsx from 'clsx';

export default function DataTable({ metrics }: { metrics: StatementMetric[] }) {
    const [localSegment, setLocalSegment] = useState<'all' | 'ip' | 'account'>('all');

    // Filter metrics based on local toggle
    const filteredLocal = useMemo(() => {
        if (localSegment === 'all') return metrics;
        if (localSegment === 'ip') return metrics.filter(m => m.statement_type.includes('IP-Level'));
        if (localSegment === 'account') return metrics.filter(m => m.statement_type.includes('Account Level'));
        return metrics;
    }, [metrics, localSegment]);

    // Aggregate ONLY by duration bucket
    const aggregatedData = useMemo(() => {
        const map = new Map<string, {
            duration_bucket: string,
            request_count: number,
            total_lead_time: number,
            count_with_lead_time: number
        }>();

        filteredLocal.forEach(m => {
            const bucket = m.duration_bucket;
            if (!map.has(bucket)) {
                map.set(bucket, {
                    duration_bucket: bucket,
                    request_count: 0,
                    total_lead_time: 0,
                    count_with_lead_time: 0
                });
            }
            const entry = map.get(bucket)!;
            entry.request_count += Number(m.request_count || 0);
            if (m.avg_lead_time) {
                entry.total_lead_time += Number(m.avg_lead_time || 0) * Number(m.request_count || 0);
                entry.count_with_lead_time += Number(m.request_count || 0);
            }
        });

        return Array.from(map.values()).map(e => ({
            ...e,
            avg_lead_time: e.count_with_lead_time > 0 ? Math.round(e.total_lead_time / e.count_with_lead_time) : 0
        })).sort((a, b) => a.duration_bucket.localeCompare(b.duration_bucket));
    }, [filteredLocal]);

    if (metrics.length === 0) return null;

    return (
        <div className="mt-8 bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-2xl">
            <div className="p-5 border-b border-slate-800 flex flex-wrap justify-between items-center gap-4 bg-slate-900/40">
                <div className="space-y-1">
                    <h3 className="font-bold text-slate-200 text-lg uppercase tracking-tight">Bucket Performance Summary</h3>
                    <p className="text-xs text-slate-500 font-medium uppercase tracking-widest">
                        {localSegment === 'all' ? 'Consolidated view of all segments & products' : `Unified totals for ${localSegment.toUpperCase()}`}
                    </p>
                </div>

                <div className="flex bg-slate-950/80 p-1 rounded-lg border border-slate-800">
                    {[
                        { id: 'all', label: 'All (Summed)' },
                        { id: 'ip', label: 'IP Level' },
                        { id: 'account', label: 'Account' }
                    ].map((s) => (
                        <button
                            key={s.id}
                            onClick={() => setLocalSegment(s.id as any)}
                            className={clsx(
                                "px-4 py-1.5 text-[10px] font-bold uppercase tracking-wider rounded-md transition-all duration-200",
                                localSegment === s.id ? "bg-sky-500 text-white shadow-lg" : "text-slate-500 hover:text-slate-300"
                            )}
                        >
                            {s.label}
                        </button>
                    ))}
                </div>
            </div>

            <div className="overflow-x-auto">
                <table className="w-full text-sm text-left border-collapse">
                    <thead className="bg-slate-950/50 text-slate-500 uppercase text-[9px] font-bold tracking-[0.2em]">
                        <tr>
                            <th className="p-5 border-b border-slate-800">Performance Category</th>
                            <th className="p-5 border-b border-slate-800 text-right">Total Requests</th>
                            <th className="p-5 border-b border-slate-800 text-right">Wtd. Avg Lead (s)</th>
                            <th className="p-5 border-b border-slate-800 text-center">Relative Impact</th>
                        </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/50">
                        {aggregatedData.map((row, i) => {
                            const totalRequests = aggregatedData.reduce((acc, r) => acc + r.request_count, 0);
                            const impactPct = totalRequests > 0 ? (row.request_count / totalRequests * 100) : 0;

                            return (
                                <tr key={i} className="hover:bg-slate-800/20 transition-colors group">
                                    <td className="p-5">
                                        <span className={clsx(
                                            "px-3 py-1.5 rounded-lg text-xs font-bold uppercase tracking-wide inline-block min-w-[160px] text-center",
                                            row.duration_bucket === "9. Over a minute" ? "bg-rose-950/40 text-rose-400 border border-rose-500/20 shadow-[0_0_15px_rgba(244,63,94,0.05)]" :
                                                ["1. less than 5 secs", "2. less than 10 secs", "3. less than 15 secs", "4. less than 20 secs"].includes(row.duration_bucket) ? "bg-emerald-950/40 text-emerald-400 border border-emerald-500/20 shadow-[0_0_15px_rgba(16,185,129,0.05)]" :
                                                    "bg-amber-950/40 text-amber-400 border border-amber-500/20 shadow-[0_0_15px_rgba(245,158,11,0.05)]"
                                        )}>
                                            {row.duration_bucket}
                                        </span>
                                    </td>
                                    <td className="p-5 text-right font-mono text-slate-100 font-bold text-lg">
                                        {row.request_count.toLocaleString()}
                                    </td>
                                    <td className="p-5 text-right font-mono text-slate-400">
                                        {row.avg_lead_time ? `${row.avg_lead_time}s` : '-'}
                                    </td>
                                    <td className="p-5">
                                        <div className="flex items-center justify-center gap-3">
                                            <div className="w-24 h-1.5 bg-slate-800 rounded-full overflow-hidden">
                                                <div
                                                    className={clsx(
                                                        "h-full rounded-full transition-all duration-500",
                                                        row.duration_bucket === "9. Over a minute" ? "bg-rose-500" :
                                                            ["1. less than 5 secs", "2. less than 10 secs", "3. less than 15 secs", "4. less than 20 secs"].includes(row.duration_bucket) ? "bg-emerald-500" :
                                                                "bg-amber-500"
                                                    )}
                                                    style={{ width: `${impactPct}%` }}
                                                />
                                            </div>
                                            <span className="text-[10px] font-bold text-slate-500 w-8">{Math.round(impactPct)}%</span>
                                        </div>
                                    </td>
                                </tr>
                            );
                        })}
                    </tbody>
                </table>
            </div>

            <div className="p-5 bg-slate-950/30 border-t border-slate-800 text-center">
                <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest">
                    Showing {aggregatedData.length} performance classifications • Total Volume: {aggregatedData.reduce((acc, r) => acc + r.request_count, 0).toLocaleString()}
                </p>
            </div>
        </div>
    );
}
