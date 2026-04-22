'use client';

import { useState } from 'react';
import { Download, Upload, Trash2, RefreshCw } from 'lucide-react';
import { exportDatabase, importDatabase, clearDatabase } from '@/lib/data';

export default function DataManagement({ onDataChange }: { onDataChange: () => void }) {
    const [loading, setLoading] = useState(false);

    const handleExport = async () => {
        setLoading(true);
        try {
            const json = await exportDatabase();
            const blob = new Blob([json], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `ci_stats_backup_${new Date().toISOString().split('T')[0]}.json`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        } catch (e) {
            console.error(e);
            alert('Export failed');
        } finally {
            setLoading(false);
        }
    };

    const handleImport = async (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if (!file) return;

        if (!confirm('This will append data to your existing database. Continue?')) return;

        setLoading(true);
        try {
            const text = await file.text();
            await importDatabase(text);
            alert('Import successful!');
            onDataChange();
        } catch (err) {
            console.error(err);
            alert('Import failed. Invalid JSON?');
        } finally {
            setLoading(false);
        }
    };

    const handleClear = async () => {
        if (!confirm('ARE YOU SURE? This will wipe all local data properly. This cannot be undone.')) return;
        setLoading(true);
        try {
            await clearDatabase();
            onDataChange();
        } finally {
            setLoading(false);
        }
    };

    return (
        <div className="bg-slate-900 border border-slate-800 rounded-xl p-4 mt-8">
            <h3 className="text-slate-200 font-semibold mb-3 flex items-center gap-2">
                <RefreshCw className="w-4 h-4" /> Data Management (Local Storage)
            </h3>
            <div className="flex flex-wrap gap-4">
                <button
                    onClick={handleExport}
                    disabled={loading}
                    className="flex items-center gap-2 px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg transition-colors text-sm"
                >
                    <Download className="w-4 h-4" /> Export Backup
                </button>

                <div className="relative">
                    <button className="flex items-center gap-2 px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg transition-colors text-sm">
                        <Upload className="w-4 h-4" /> Import Backup
                    </button>
                    <input
                        type="file"
                        accept=".json"
                        onChange={handleImport}
                        className="absolute inset-0 opacity-0 cursor-pointer"
                        disabled={loading}
                    />
                </div>

                <button
                    onClick={handleClear}
                    disabled={loading}
                    className="flex items-center gap-2 px-4 py-2 bg-red-900/30 hover:bg-red-900/50 text-red-200 border border-red-900/50 rounded-lg transition-colors text-sm ml-auto"
                >
                    <Trash2 className="w-4 h-4" /> Reset Database
                </button>
            </div>
        </div>
    );
}
