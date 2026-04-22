export interface StatementMetric {
    id?: number | string;
    created_at?: string;
    report_start_date: string;
    report_end_date: string;
    statement_type: string;
    product_type: string;
    duration_bucket: string;
    request_count: number;
    avg_lead_time?: number;
    p95_time?: number;
}

export type DurationBucket =
    | "1. less than 5 secs"
    | "2. less than 10 secs"
    | "3. less than 15 secs"
    | "4. less than 20 secs"
    | "5. less than 25 secs"
    | "6. less than 30 secs"
    | "7. less than 45 secs"
    | "8. less than 1 min"
    | "9. Over a minute";

export const DURATION_BUCKETS: DurationBucket[] = [
    "1. less than 5 secs",
    "2. less than 10 secs",
    "3. less than 15 secs",
    "4. less than 20 secs",
    "5. less than 25 secs",
    "6. less than 30 secs",
    "7. less than 45 secs",
    "8. less than 1 min",
    "9. Over a minute"
];
