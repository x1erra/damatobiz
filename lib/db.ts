
import Dexie, { type EntityTable } from 'dexie';
import { StatementMetric } from '@/types';

const db = new Dexie('CIStatsDB') as Dexie & {
    metrics: EntityTable<StatementMetric, 'id'> & { id: number };
};

// Schema definition
// We index fields that we want to query or sort by
db.version(1).stores({
    metrics: '++id, report_start_date, statement_type, [report_start_date+report_end_date+statement_type+product_type+duration_bucket]'
});

export { db };
