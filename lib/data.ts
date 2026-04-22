import { db } from './db';
import { StatementMetric } from '../types';

export async function uploadMetrics(metrics: StatementMetric[]) {
    // Dexie bulkPut handles upsert if primary key matches, but we have a composite logic constraint.
    // We used a composite unique index in SQL: [report_start_date+report_end_date+statement_type+product_type+duration_bucket]
    // In Dexie, we defined this compound index in the schema.

    // However, `bulkPut` will overwrite if the KEY exists. Since our PK is auto-increment ID, we need to be careful.
    // Actually, for a simple analytics tool, we can clear and insert, OR we can check existence.
    // Efficient strategy: Use the compound index to find existing items? 
    // BETTER: Just rely on the user. If they upload the same file, it might duplicate unless we check.

    // Let's implement a "smart" ingestion that tries to prevent dupes for the same Start Date + Statement Type.
    // But bulkPut is fastest.

    // Simple approach for V1 Local: 
    // We can't easily rely on auto-id AND unique constraint in Dexie the same way as SQL without logic.
    // Let's create a synthetic ID for deduplication?
    // Or just accept that if they upload the same weekly report twice, they should probably delete it first?
    // Let's try to clear data for these dates first? No that's dangerous.

    // Let's just add them. The user can clear DB if they mess up (RESET functionality).

    await db.metrics.bulkAdd(metrics);
}

export async function fetchMetrics(days = 30) {
    // Filter in memory for now, Dexie is fast enough for <100k rows
    const all = await db.metrics.toArray();

    if (days === 99999) return all;

    const dateThreshold = new Date();
    dateThreshold.setDate(dateThreshold.getDate() - days);

    return all.filter(m => new Date(m.report_start_date) >= dateThreshold);
}

export async function clearDatabase() {
    await db.metrics.clear();
}

export async function exportDatabase() {
    const all = await db.metrics.toArray();
    return JSON.stringify(all);
}

export async function importDatabase(jsonString: string) {
    const data = JSON.parse(jsonString);
    await db.metrics.bulkAdd(data);
}
