'use client';

import { useState } from 'react';
import Papa from 'papaparse';
import { UploadCloud, CheckCircle, AlertCircle } from 'lucide-react';
import { uploadMetrics } from '@/lib/data';
import { DURATION_BUCKETS, StatementMetric } from '@/types';

export default function UploadZone({ onUploadComplete }: { onUploadComplete: () => void }) {
    const [isUploading, setIsUploading] = useState(false);
    const [status, setStatus] = useState<'idle' | 'success' | 'error'>('idle');
    const [message, setMessage] = useState('');

    const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
        const files = e.target.files;
        if (!files || files.length === 0) return;

        setIsUploading(true);
        setStatus('idle');
        setMessage('');

        try {
            let totalRows = 0;

            for (let i = 0; i < files.length; i++) {
                const file = files[i];

                const parsedData = await new Promise<StatementMetric[]>((resolve, reject) => {
                    Papa.parse(file, {
                        header: false,
                        skipEmptyLines: true,
                        complete: (results) => {
                            const rows = results.data as string[][];
                            const metrics: StatementMetric[] = [];

                            // Helper to normalize date: "20260101" -> "2026-01-01"
                            const normalizeDate = (d: string) => {
                                if (!d) return "";
                                const clean = d.replace('=', '').replace(/"/g, '').trim();
                                if (/^\d{8}$/.test(clean)) {
                                    return `${clean.substring(0, 4)}-${clean.substring(4, 6)}-${clean.substring(6, 8)}`;
                                }
                                return clean;
                            };

                            let currentStartDate = "";
                            let currentEndDate = "";
                            let lastStatementType = "Unknown";
                            let lastProductType = "N/A";

                            rows.forEach(row => {
                                // Extract Dates from metadata lines
                                if (row[0]?.includes('Start Date')) currentStartDate = normalizeDate(row[1]);
                                if (row[0]?.includes('End Date')) currentEndDate = normalizeDate(row[1]);

                                // Forward Fill Logic for segments
                                const currentStatementType = row[2]?.trim();
                                const currentProductType = row[3]?.trim();

                                if (currentStatementType && currentStatementType !== 'Statement Type') lastStatementType = currentStatementType;
                                if (currentProductType && currentProductType !== 'Product Type') lastProductType = currentProductType;

                                const duration = row[4]?.trim();

                                if (DURATION_BUCKETS.includes(duration as any)) {
                                    metrics.push({
                                        report_start_date: currentStartDate || row[0],
                                        report_end_date: currentEndDate || row[1],
                                        statement_type: lastStatementType,
                                        product_type: lastProductType,
                                        duration_bucket: duration as any,
                                        request_count: parseInt(row[5]) || 0,
                                        avg_lead_time: parseFloat(row[7]) || 0,
                                        p95_time: parseFloat(row[9]) || 0
                                    });
                                }
                            });
                            resolve(metrics);
                        },
                        error: (error) => reject(error)
                    });
                });

                if (parsedData.length > 0) {
                    await uploadMetrics(parsedData);
                    totalRows += parsedData.length;
                }
            }

            setStatus('success');
            setMessage(`Successfully uploaded ${totalRows} metrics!`);
            onUploadComplete();
        } catch (err: any) {
            console.error(err);
            setStatus('error');
            setMessage(err.message || 'Failed to parse/upload files.');
        } finally {
            setIsUploading(false);
        }
    };

    return (
        <div className="mb-8 p-6 border-2 border-dashed border-slate-700 hover:border-sky-500 rounded-xl bg-slate-900 transition-colors text-center cursor-pointer relative group">
            <input
                type="file"
                multiple
                accept=".csv"
                onChange={handleFileUpload}
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                disabled={isUploading}
            />

            <div className="flex flex-col items-center justify-center space-y-4">
                {isUploading ? (
                    <div className="animate-spin w-10 h-10 border-4 border-sky-500 border-t-transparent rounded-full" />
                ) : status === 'success' ? (
                    <CheckCircle className="w-12 h-12 text-green-500" />
                ) : status === 'error' ? (
                    <AlertCircle className="w-12 h-12 text-red-500" />
                ) : (
                    <UploadCloud className="w-12 h-12 text-slate-400 group-hover:text-sky-500 transition-colors" />
                )}

                <div>
                    <h3 className="text-lg font-semibold text-slate-200">
                        {isUploading ? 'Processing...' : 'Upload Weekly Reports'}
                    </h3>
                    <p className="text-sm text-slate-400 mt-1">
                        {message || 'Drag & drop CSV files or click to browse'}
                    </p>
                </div>
            </div>
        </div>
    );
}
