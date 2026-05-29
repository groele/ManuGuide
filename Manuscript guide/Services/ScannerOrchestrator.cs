using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Manuscript_guide.Models;
using Manuscript_guide.Scanners;

namespace Manuscript_guide.Services
{
    public static class ScannerOrchestrator
    {
        public static async System.Threading.Tasks.Task<List<IssueItem>> ScanAsync(Document doc, string moduleType, Action<int> progressCallback)
        {
            // 1. Build document snapshot in Word main thread (STA)
            progressCallback?.Invoke(10);
            var snapshot = DocumentSnapshotBuilder.Build(doc);
            progressCallback?.Invoke(30);

            // 2. Offload actual scanning to background thread
            return await System.Threading.Tasks.Task.Run(() =>
            {
                var allIssues = new List<IssueItem>();

                // Set up the scan context using the snapshot
                using (DocumentScanContext.Begin(doc, snapshot))
                {
                    progressCallback?.Invoke(40);

                    // Load matching specialized scanners
                    var scannersToRun = new List<ISpecializedScanner>();
                    if (string.Equals(moduleType, "all", StringComparison.OrdinalIgnoreCase))
                    {
                        foreach (var module in DiagnosticModuleRegistry.FullScanModules)
                        {
                            var scanner = module.CreateScanner();
                            if (scanner != null)
                            {
                                scannersToRun.Add(scanner);
                            }
                        }
                    }
                    else
                    {
                        var scanner = DiagnosticModuleRegistry.CreateScanner(moduleType);
                        if (scanner != null)
                        {
                            scannersToRun.Add(scanner);
                        }
                    }

                    if (scannersToRun.Count == 0)
                    {
                        return allIssues;
                    }

                    // Sequential scanning on this thread to preserve [ThreadStatic] DocumentScanContext.
                    // Parallel.ForEach would spawn worker threads that cannot see the context,
                    // because current is [ThreadStatic] and each worker thread gets its own null copy.
                    // Sequential is fine since scanners are pure in-memory regex operations on the snapshot.
                    int completed = 0;

                    foreach (var scanner in scannersToRun)
                    {
                        try
                        {
                            // doc is passed as null to ensure no accidental COM calls on background thread
                            var scannerIssues = scanner.Scan(null); 
                            if (scannerIssues != null && scannerIssues.Count > 0)
                            {
                                var enriched = IssueMetadataService.EnrichAll(scannerIssues);
                                allIssues.AddRange(enriched);
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error scanning with {scanner.ModuleType}: " + ex.Message);
                        }

                        completed++;
                        int progress = 40 + (completed * 50 / scannersToRun.Count);
                        progressCallback?.Invoke(progress);
                    }
                }

                progressCallback?.Invoke(95);
                return allIssues;
            });
        }
    }
}
