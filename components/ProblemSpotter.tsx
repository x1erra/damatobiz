'use client';

import { StatementMetric } from '@/types';
import { calculateHealthScore } from '@/lib/analytics';
import { TrendingDown, Package, Activity, Calendar, HelpCircle, Layers } from 'lucide-react';
import { useMemo } from 'react';

function Tooltip({ text }: { text: string }) {
    if (!text) return null;
    return (
        <div className="absolute bottom-full left-1/2 -translate-x-1/2 mb-2 p-2 bg-slate-800 text-slate-200 text-[11px] rounded shadow-xl border border-slate-700 w-48 invisible group-hover:visible z-[100] transition-all opacity-0 group-hover:opacity-100 backdrop-blur-sm pointer-events-none leading-relaxed">
            {text}
            <div className="absolute top-full left-1/2 -translate-x-1/2 border-8 border-transparent border-t-slate-800" />
        </div>
    );
}

export default function ProblemSpotter({ metrics }: { metrics: StatementMetric[] }) {
    const insights = useMemo(() => {
        if (metrics.length === 0) return null;

        const reportDates = Array.from(new Set(metrics.map(m => m.report_start_date))).sort();

        // 1. Seasonal Trends
        const quarters = [
            { name: 'Q1', months: [0, 1, 2], total: 0, fast: 0 },
            { name: 'Q2', months: [3, 4, 5], total: 0, fast: 0 },
            { name: 'Q3', months: [6, 7, 8], total: 0, fast: 0 },
            { name: 'Q4', months: [9, 10, 11], total: 0, fast: 0 }
        ];
        metrics.forEach(m => {
            const date = new Date(m.report_start_date);
            const q = quarters.find(q => q.months.includes(date.getMonth()));
            if (q) {
                const count = Number(m.request_count || 0);
                q.total += count;
                if (["1. less than 5 secs", "2. less than 10 secs", "3. less than 15 secs", "4. less than 20 secs"].includes(m.duration_bucket)) q.fast += count;
            }
        });
        const worstQuarter = quarters.map(q => ({ ...q, score: q.total > 0 ? (q.fast / q.total) * 100 : 100 })).filter(q => q.total > 0).sort((a, b) => a.score - b.score)[0];

        // 2. Product Bottleneck
        const products = Array.from(new Set(metrics.map(m => m.product_type)));
        const productStats = products.map(p => {
            const pMetrics = metrics.filter(m => m.product_type === p);
            const total = pMetrics.reduce((acc, m) => acc + Number(m.request_count || 0), 0);
            const critical = pMetrics.filter(m => m.duration_bucket === "9. Over a minute").reduce((acc, m) => acc + Number(m.request_count || 0), 0);
            return { name: p, criticalCount: critical, criticalPct: total > 0 ? (critical / total) * 100 : 0 };
        }).sort((a, b) => b.criticalCount - a.criticalCount);

        // 3. Segment Analysis (IP vs Account)
        const ipMetrics = metrics.filter(m => m.statement_type.includes('IP-Level'));
        const accMetrics = metrics.filter(m => m.statement_type.includes('Account Level'));
        const ipHealth = calculateHealthScore(ipMetrics);
        const accHealth = calculateHealthScore(accMetrics);
        const segmentGap = Math.abs(ipHealth - accHealth);
        const worstSegment = ipHealth < accHealth ? { name: 'IP-Level', score: ipHealth } : { name: 'Account Level', score: accHealth };

        // 4. Performance Dips
        const weeklyData = reportDates.map(date => {
            const weekMetrics = metrics.filter(m => m.report_start_date === date);
            const total = weekMetrics.reduce((acc, m) => acc + Number(m.request_count || 0), 0);
            const fast = weekMetrics.filter(m => ["1. less than 5 secs", "2. less than 10 secs", "3. less than 15 secs", "4. less than 20 secs"].includes(m.duration_bucket)).reduce((acc, m) => acc + Number(m.request_count || 0), 0);
            return { date, score: total > 0 ? (fast / total) * 100 : 100 };
        });
        let biggestDip = { from: '', to: '', drop: 0 };
        for (let i = 1; i < weeklyData.length; i++) {
            const drop = weeklyData[i - 1].score - weeklyData[i].score;
            if (drop > biggestDip.drop) biggestDip = { from: weeklyData[i - 1].date, to: weeklyData[i].date, drop };
        }

        return { worstQuarter, productStats, worstSegment, segmentGap, biggestDip };
    }, [metrics]);

    if (!insights) return null;
    const { worstQuarter, productStats, worstSegment, segmentGap, biggestDip } = insights;
    const primaryBottleneck = productStats[0];

    return (
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-8">
            <InsightCard title="Primary Impact Item" icon={<Package className="text-sky-400 w-4 h-4" />} tooltip="The product generating the largest volume of critical wait times.">
                <span className="font-bold text-slate-200">{primaryBottleneck?.name}</span> is the biggest driver of critical wait with <span className="font-bold text-sky-400">{primaryBottleneck?.criticalCount.toLocaleString()}</span> requests affected.
            </InsightCard>

            <InsightCard title="Statement Type Comparison" icon={<Layers className="text-sky-400 w-4 h-4" />} tooltip="Compares performance between IP-Level and Account Level reporting types.">
                {segmentGap > 5 ? (
                    <>Performance is <span className="text-rose-400 font-bold">unbalanced</span>. <span className="font-bold">{worstSegment.name}</span> is lagging behind by <span className="font-bold text-rose-400">{segmentGap.toFixed(1)}%</span>.</>
                ) : (
                    <>Health scores are <span className="text-emerald-400 font-medium">consistent</span> across both reporting segments.</>
                )}
            </InsightCard>

            <InsightCard title="Seasonal Trend" icon={<Calendar className="text-sky-400 w-4 h-4" />} tooltip="Identifies the quarter with the lowest average health score.">
                {worstQuarter ? <><span className="font-bold">{worstQuarter.name}</span> typically sees the lowest performance with an average health of <span className="font-bold text-rose-400">{worstQuarter.score.toFixed(1)}%</span>.</> : 'Calculating...'}
            </InsightCard>

            <InsightCard title="Volatility Alert" icon={<TrendingDown className="text-sky-400 w-4 h-4" />} tooltip="Highlights the sharpest week-over-week drop in performance.">
                {biggestDip.drop > 2 ? <>A <span className="font-bold text-rose-400">{biggestDip.drop.toFixed(1)}%</span> performance drop occurred between {biggestDip.from} and {biggestDip.to}.</> : <>Performance is currently <span className="text-emerald-400 font-medium">highly stable</span> (&lt; 2% variance).</>}
            </InsightCard>
        </div>
    );
}

function InsightCard({ title, icon, children, tooltip }: { title: string, icon: React.ReactNode, children: React.ReactNode, tooltip: string }) {
    return (
        <div className="bg-slate-900/40 border border-slate-800/60 p-4 rounded-xl relative group hover:bg-slate-900/60 transition-all cursor-default">
            <div className="flex items-center gap-2 mb-2">
                {icon}
                <h4 className="font-semibold text-slate-300 text-[10px] uppercase tracking-widest">{title}</h4>
                <HelpCircle className="w-3 h-3 text-slate-700 ml-auto" />
            </div>
            <p className="text-sm text-slate-400 leading-relaxed">{children}</p>
            <Tooltip text={tooltip} />
        </div>
    );
}
