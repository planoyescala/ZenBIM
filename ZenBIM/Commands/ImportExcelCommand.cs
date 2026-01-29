//-----------------------------------------------------------------------------------------
// <copyright file="ImportExcelCommand.cs" company="plano y escala">
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
using ZenBIM.Views;
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace ZenBIM.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ImportExcelCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // 1. UI Initialization
            var ui = new ImportExcelView();

            if (ui.ShowDialog() != true) return Result.Cancelled;

            string safePath = ui.SelectedFilePath;
            string safeSheet = ui.SelectedSheetName;
            string safeRange = ui.SelectedRangeName;

            // SAFETY CHECK
            if (IsFileLocked(safePath))
            {
                Autodesk.Revit.UI.TaskDialog.Show("ZenBIM Error", "The file is locked by another process. Import cancelled.");
                return Result.Cancelled;
            }

            // 2. Progress Feedback
            var progress = new ZenProgressView("ZenBIM: Excel Sync", "Processing spreadsheet data...");
            progress.Show();

            ViewSchedule? scheduleCreated = null;

            using (Transaction t = new Transaction(doc, "ZenBIM: Import Excel"))
            {
                t.Start();
                try
                {
                    scheduleCreated = RunImportProcess(doc, safePath, safeSheet, safeRange, progress);
                    t.Commit();
                }
                catch (Exception ex)
                {
                    if (t.HasStarted()) t.RollBack();
                    progress.Close();
                    Autodesk.Revit.UI.TaskDialog.Show("ZenBIM Error", "Critical process error: " + ex.Message);
                    return Result.Failed;
                }
            }

            progress.Close();
            if (scheduleCreated != null) uidoc.ActiveView = scheduleCreated;

            return Result.Succeeded;
        }

        private bool IsFileLocked(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath)) return false;
            try
            {
                using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    stream.Close();
                }
            }
            catch (IOException) { return true; }
            return false;
        }

        private ViewSchedule RunImportProcess(Document doc, string path, string sheetName, string rangeName, ZenProgressView p)
        {
            using (var workbook = new XLWorkbook(path))
            {
                IXLRange? rng = null;

                // Get Worksheet explicitly
                var worksheet = workbook.Worksheet(sheetName);
                if (worksheet == null) throw new InvalidOperationException($"Worksheet '{sheetName}' not found.");

                // CORRECCIÓN CRÍTICA:
                // No usar .DefinedName("nombre") directamente porque lanza excepción si no existe en ese ámbito.
                // Usamos LINQ para buscar de forma segura.

                // 1. Try to find Sheet Scoped named range safely
                var sheetDefinedName = worksheet.DefinedNames.FirstOrDefault(n => n.Name == rangeName);
                if (sheetDefinedName != null)
                {
                    rng = sheetDefinedName.Ranges.FirstOrDefault();
                }

                // 2. Fallback to Global defined name safely
                if (rng == null)
                {
                    var globalDefinedName = workbook.DefinedNames.FirstOrDefault(n => n.Name == rangeName);
                    rng = globalDefinedName?.Ranges.FirstOrDefault();
                }

                if (rng == null) throw new InvalidOperationException($"Range '{rangeName}' could not be resolved in Sheet or Workbook scope.");

                int rows = rng.RowCount();
                int cols = rng.ColumnCount();
                int startRow = rng.FirstRow().RowNumber();
                int startCol = rng.FirstColumn().ColumnNumber();

                // 3. Create Schedule
                ViewSchedule schedule = ViewSchedule.CreateSchedule(doc, new ElementId(BuiltInCategory.OST_GenericModel));

                // Naming logic
                string dateStr = DateTime.Now.ToString("yyyy-MM-dd");
                string baseName = $"ZenBIM_{dateStr}_{sheetName}_{rangeName}";
                string finalName = baseName;
                int nameCount = 1;

                var existingNames = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSchedule))
                    .Cast<ViewSchedule>()
                    .Select(v => v.Name ?? string.Empty)
                    .ToList();

                while (existingNames.Contains(finalName))
                {
                    nameCount++;
                    finalName = $"{baseName} ({nameCount})";
                }
                schedule.Name = finalName;

                // 4. Setup Schedule Structure
                var field = schedule.Definition.GetSchedulableFields()
                    .FirstOrDefault(f => f.GetName(doc).Contains("Comments") || f.GetName(doc).Contains("Comentarios"));

                if (field != null)
                {
                    ScheduleField sf = schedule.Definition.AddField(field);
                    schedule.Definition.AddFilter(new ScheduleFilter(sf.FieldId, ScheduleFilterType.Equal, "ZENBIM_HIDDEN_DATA"));
                }

                schedule.Definition.ShowTitle = true;
                schedule.Definition.ShowHeaders = false;

                TableSectionData header = schedule.GetTableData().GetSectionData(SectionType.Header);
                TableSectionData body = schedule.GetTableData().GetSectionData(SectionType.Body);

                // Scaling Constants
                double baseSizeFeet = 0.00590551;
                double fontFactor = (baseSizeFeet / 7.0) * 1125.0;
                double colFactor = 0.0656168 / 10.71;
                double rowFactor = 0.0229659 / 15.75;

                // Sync Width
                double totalWidth = 0;
                for (int c = 0; c < cols; c++) totalWidth += (worksheet.Column(startCol + c).Width * colFactor);
                try { body.SetColumnWidth(0, totalWidth); } catch { }

                // Build Grid
                for (int c = 1; c < cols; c++) header.InsertColumn(c);
                for (int r = 0; r < rows; r++) header.InsertRow(header.NumberOfRows);

                // Populate Data
                for (int r = 0; r < rows; r++)
                {
                    p.Update(((double)(r + 1) / rows) * 100);
                    int revitRow = r + 1;

                    header.SetRowHeight(revitRow, worksheet.Row(startRow + r).Height * rowFactor);

                    for (int c = 0; c < cols; c++)
                    {
                        var cell = worksheet.Cell(startRow + r, startCol + c);
                        if (r == 0) header.SetColumnWidth(c, worksheet.Column(startCol + c).Width * colFactor);

                        string addr = cell.Address.ToString();
                        // Safe cast for MergedRanges
                        IXLRange? merged = worksheet.MergedRanges.FirstOrDefault(m => m.Contains(addr)) as IXLRange;

                        if (merged != null && merged.FirstCell().Address.ToString() == addr)
                        {
                            try { header.MergeCells(new TableMergedCell(revitRow, c, revitRow + merged.RowCount() - 1, c + merged.ColumnCount() - 1)); } catch { }
                        }

                        header.SetCellText(revitRow, c, cell.GetFormattedString() ?? string.Empty);

                        // Styling
                        TableCellStyle s = header.GetTableCellStyle(revitRow, c);
                        TableCellStyleOverrideOptions o = s.GetCellStyleOverrideOptions();

                        s.FontName = cell.Style.Font.FontName ?? "Arial";
                        s.TextSize = Math.Max(cell.Style.Font.FontSize * fontFactor, baseSizeFeet * 1125.0);
                        o.FontSize = true;
                        if (cell.Style.Font.Bold) { s.IsFontBold = true; o.Bold = true; }

                        // Color Logic
                        var xlFill = cell.Style.Fill;
                        System.Drawing.Color? finalColor = GetSafeColor(xlFill);

                        if (finalColor.HasValue)
                        {
                            s.BackgroundColor = new Autodesk.Revit.DB.Color(finalColor.Value.R, finalColor.Value.G, finalColor.Value.B);
                            o.BackgroundColor = true;
                        }

                        var align = cell.Style.Alignment.Horizontal;
                        s.FontHorizontalAlignment = align == XLAlignmentHorizontalValues.Center ? HorizontalAlignmentStyle.Center :
                                                     align == XLAlignmentHorizontalValues.Right ? HorizontalAlignmentStyle.Right :
                                                     HorizontalAlignmentStyle.Left;

                        s.SetCellStyleOverrideOptions(o);
                        header.SetCellStyle(revitRow, c, s);
                    }
                }

                try { header.RemoveRow(0); } catch { }

                return schedule;
            }
        }

        private System.Drawing.Color? GetSafeColor(IXLFill fill)
        {
            try
            {
                if (fill.BackgroundColor.ColorType == XLColorType.Theme)
                    return GetThemeFallback(fill.BackgroundColor.ThemeColor);

                if (!fill.BackgroundColor.Equals(XLColor.NoColor))
                    return GetColorOrApproximate(fill.BackgroundColor);

                if (!fill.PatternColor.Equals(XLColor.NoColor))
                    return GetColorOrApproximate(fill.PatternColor);
            }
            catch { }
            return null;
        }

        private System.Drawing.Color? GetColorOrApproximate(XLColor xlColor)
        {
            try
            {
                System.Drawing.Color c = xlColor.Color;
                if (c.R == 255 && c.G == 255 && c.B == 255) return null;
                return c;
            }
            catch
            {
                return GetThemeFallback(xlColor.ThemeColor);
            }
        }

        private System.Drawing.Color? GetThemeFallback(XLThemeColor theme)
        {
            switch (theme)
            {
                case XLThemeColor.Accent1: return System.Drawing.Color.FromArgb(189, 215, 238);
                case XLThemeColor.Accent2: return System.Drawing.Color.FromArgb(250, 192, 144);
                case XLThemeColor.Accent3: return System.Drawing.Color.FromArgb(219, 219, 219);
                case XLThemeColor.Accent4: return System.Drawing.Color.FromArgb(255, 230, 153);
                case XLThemeColor.Accent5: return System.Drawing.Color.FromArgb(157, 195, 230);
                case XLThemeColor.Accent6: return System.Drawing.Color.FromArgb(169, 208, 142);
                case XLThemeColor.Text1: return System.Drawing.Color.FromArgb(240, 240, 240);
                case XLThemeColor.Text2: return null;
                default: return System.Drawing.Color.FromArgb(242, 242, 242);
            }
        }
    }
}