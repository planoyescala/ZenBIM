//-----------------------------------------------------------------------------------------
// <copyright file="BatchPrintCommand.cs" company="plano y escala">
// Copyright (c) 2026 plano y escala.
//
// ZenBIM is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// </copyright>
// <author>plano y escala</author>
//-----------------------------------------------------------------------------------------

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using ZenBIM.Views;

namespace ZenBIM.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class BatchPrintCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            var ui = new BatchPrintView(doc);
            if (ui.ShowDialog() != true) return Result.Cancelled;

            // UI Parameters
            List<ViewSheet> sheetsToExport = ui.SelectedSheets;
            string outputFolder = ui.OutputFolder;
            string namingRule = ui.NamingRule;
            bool combine = ui.CombineFiles;
            int exportMode = ui.ExportMode; // 0=PDF, 1=DWG, 2=Both
            string dwgSetup = ui.SelectedDWGSetup;

            // Printer Icon
            string printerIconGeometry = "M18,3H6V7H18M19,12A1,1 0 0,1 18,11A1,1 0 0,1 19,10A1,1 0 0,1 20,11A1,1 0 0,1 19,12M16,19H8V14H16M19,8H5A3,3 0 0,0 2,11V17H6V21H18V17H22V11A3,3 0 0,0 19,8Z";

            var progress = new ZenProgressView("ZenBIM: Printing", "Initializing export...");
            progress.SetIcon(printerIconGeometry);
            progress.Show();

            try
            {
                // Prepare Options
                PDFExportOptions pdfOptions = new PDFExportOptions
                {
                    Combine = combine,
                    ColorDepth = ui.SelectedColor,
                    RasterQuality = ui.SelectedQuality,
                    StopOnError = false,
                    HideScopeBoxes = ui.HideScopeBoxes,
                    HideReferencePlane = true
                };

                DWGExportOptions dwgOptions = new DWGExportOptions();
                if (exportMode == 1 || exportMode == 2)
                {
                    var setups = new FilteredElementCollector(doc).OfClass(typeof(ExportDWGSettings)).Cast<ExportDWGSettings>();
                    var setup = setups.FirstOrDefault(x => x.Name == dwgSetup);
                    if (setup != null) dwgOptions = setup.GetDWGExportOptions();
                    dwgOptions.MergedViews = true;
                }

                int total = sheetsToExport.Count;
                int current = 0;

                // --- COMBINED PDF CASE (Special case, single transaction usually fine) ---
                if (combine && (exportMode == 0 || exportMode == 2))
                {
                    using (Transaction t = new Transaction(doc, "ZenBIM Combined Export"))
                    {
                        t.Start();
                        progress.Update(50);
                        string finalName = Sanitize(ParseTokens(namingRule, null, doc));
                        if (string.IsNullOrWhiteSpace(finalName)) finalName = $"Batch_{DateTime.Now:MMdd}";

                        pdfOptions.FileName = finalName;

                        // Handle collision
                        string targetPath = Path.Combine(outputFolder, finalName + ".pdf");
                        if (File.Exists(targetPath)) File.Delete(targetPath);

                        doc.Export(outputFolder, sheetsToExport.Select(s => s.Id).ToList(), pdfOptions);
                        t.Commit();
                    }
                    // If combined PDF was the only task, we are done with PDF part.
                    // But if DWG is also selected, we continue below.
                }

                // --- INDIVIDUAL FILES LOOP (Transaction PER SHEET) ---
                // We only enter here if NOT combining PDF, or if we need to do DWGs
                bool doIndividualPdf = (exportMode == 0 || exportMode == 2) && !combine;
                bool doDwg = (exportMode == 1 || exportMode == 2);

                if (doIndividualPdf || doDwg)
                {
                    foreach (ViewSheet sheet in sheetsToExport)
                    {
                        current++;
                        double progressVal = (double)current / total * 100;
                        progress.Update(progressVal);
                        progress.TxtSubStatus.Text = $"Processing: {sheet.SheetNumber}";

                        // Generate unique Temp ID for this specific sheet export
                        string tempId = "ZEN_" + Guid.NewGuid().ToString("N").Substring(0, 8);

                        // Calculate Final Name based on Rule
                        string finalBaseName = Sanitize(ParseTokens(namingRule, sheet, doc));

                        // START TRANSACTION FOR THIS SHEET
                        using (Transaction tx = new Transaction(doc, "ZenBIM Single Export"))
                        {
                            tx.Start();

                            // Export PDF (Individual)
                            if (doIndividualPdf)
                            {
                                pdfOptions.FileName = tempId;
                                doc.Export(outputFolder, new List<ElementId> { sheet.Id }, pdfOptions);
                            }

                            // Export DWG
                            if (doDwg)
                            {
                                // DWG Export uses 'tempId' as prefix
                                doc.Export(outputFolder, tempId, new List<ElementId> { sheet.Id }, dwgOptions);
                            }

                            tx.Commit();
                        } // Transaction Closed -> Files should be flushed

                        // POST-PROCESS RENAMING (Outside Transaction)
                        if (doIndividualPdf)
                            RenameWithOriginalLogic(outputFolder, tempId, sheet, finalBaseName + ".pdf", "*.pdf");

                        if (doDwg)
                            RenameWithOriginalLogic(outputFolder, tempId, sheet, finalBaseName + ".dwg", "*.dwg");
                    }
                }
            }
            catch (Exception ex)
            {
                progress.Close();
                Autodesk.Revit.UI.TaskDialog.Show("Error", "Export Failed:\n" + ex.Message);
                return Result.Failed;
            }
            finally
            {
                progress.Close();
            }

            if (ui.OpenAfter) System.Diagnostics.Process.Start("explorer.exe", outputFolder);

            return Result.Succeeded;
        }

        // Adapted from your working code
        private void RenameWithOriginalLogic(string folder, string tempId, ViewSheet sheet, string finalName, string extensionPattern)
        {
            // Initial Wait - Critical for file system catch-up
            System.Threading.Thread.Sleep(1000);

            var directory = new DirectoryInfo(folder);
            if (!directory.Exists) return;

            // Fuzzy Search: Find file containing TempID OR SheetNumber OR SheetName
            // Revit sometimes ignores the prefix and uses View Name, so we look for any trace.
            // We prioritize TempId if found.
            var fileToRename = directory.GetFiles(extensionPattern)
                .Where(f => f.Name.Contains(tempId) ||
                            f.Name.Contains(sheet.SheetNumber) ||
                            (sheet.Name != null && f.Name.Contains(sheet.Name)))
                .OrderByDescending(f => f.LastWriteTime) // Get the most recent one
                .FirstOrDefault();

            if (fileToRename != null)
            {
                string finalPath = Path.Combine(folder, finalName);

                // Ignore if source and target are already the same
                if (fileToRename.FullName.Equals(finalPath, StringComparison.OrdinalIgnoreCase)) return;

                int attempts = 0;
                while (attempts < 3)
                {
                    try
                    {
                        if (File.Exists(finalPath)) File.Delete(finalPath);
                        File.Move(fileToRename.FullName, finalPath);
                        break; // Success
                    }
                    catch (IOException)
                    {
                        // File locked, wait and retry
                        System.Threading.Thread.Sleep(1500);
                        attempts++;
                    }
                }
            }
        }

        private string ParseTokens(string rule, ViewSheet? sheet, Document doc)
        {
            if (string.IsNullOrEmpty(rule)) return "Unnamed";

            return Regex.Replace(rule, @"\{(.*?)\}", match =>
            {
                string paramName = match.Groups[1].Value;

                if (sheet != null)
                {
                    Parameter p = sheet.LookupParameter(paramName);
                    if (p != null) return p.AsString() ?? p.AsValueString() ?? "";
                    if (paramName == "Sheet Number") return sheet.SheetNumber;
                    if (paramName == "Sheet Name") return sheet.Name;
                }

                if (doc.ProjectInformation != null)
                {
                    Parameter pInfo = doc.ProjectInformation.LookupParameter(paramName);
                    if (pInfo != null) return pInfo.AsString() ?? pInfo.AsValueString() ?? "";
                    if (paramName == "Project Number") return doc.ProjectInformation.Number;
                    if (paramName == "Project Name") return doc.ProjectInformation.Name;
                }

                return "Unknown";
            });
        }

        private string Sanitize(string name)
        {
            if (string.IsNullOrEmpty(name)) return "Unnamed";
            foreach (char c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
            return name;
        }
    }
}