import { StatementMetric, DURATION_BUCKETS } from '../types';

export function calculateHealthScore(metrics: StatementMetric[]): number {
    if (metrics.length === 0) return 0;

    const total = metrics.reduce((acc, m) => acc + Number(m.request_count || 0), 0);
    const fast = metrics
        .filter(m => [
            "1. less than 5 secs",
            "2. less than 10 secs",
            "3. less than 15 secs",
            "4. less than 20 secs"
        ].includes(m.duration_bucket))
        .reduce((acc, m) => acc + Number(m.request_count || 0), 0);

    return total > 0 ? (fast / total) * 100 : 0;
}

export function aggregateByDate(metrics: StatementMetric[]) {
    const grouped = new Map<string, StatementMetric[]>();

    metrics.forEach(m => {
        const list = grouped.get(m.report_start_date) || [];
        list.push(m);
        grouped.set(m.report_start_date, list);
    });

    return Array.from(grouped.entries())
        .sort((a, b) => new Date(a[0]).getTime() - new Date(b[0]).getTime())
        .map(([date, items]) => ({
            date,
            healthScore: calculateHealthScore(items),
            criticalCount: items
                .filter(m => m.duration_bucket === "9. Over a minute")
                .reduce((acc, m) => acc + Number(m.request_count || 0), 0),
            totalRequests: items.reduce((acc, m) => acc + Number(m.request_count || 0), 0),
            metrics: items
        }));
}
