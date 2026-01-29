//-----------------------------------------------------------------------------------------
// <copyright file="ExportExcelView.xaml.cs" company="plano y escala">
// Copyright (c) 2026 plano y escala.
//
// ZenBIM is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// </copyright>
// <author>plano y escala</author>
//-----------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Diagnostics; // Required for Process.Start
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;
using ZenBIM.Core;
using ClosedXML.Excel; // Using ClosedXML library

namespace ZenBIM.Views
{
    public partial class ExportExcelView : Window
    {
        private readonly Document _doc;
        private List<ScheduleItem> _allSchedules = new List<ScheduleItem>();
        private string _destinationPath = string.Empty;

        public ExportExcelView(Document doc)
        {
            InitializeComponent();
            _doc = doc;
            LoadSchedules();
        }

        private void LoadSchedules()
        {
            _allSchedules = new List<ScheduleItem>();

            // Collect all ViewSchedule elements that are not templates
            var collector = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(v => !v.IsTemplate && v.ViewType == ViewType.Schedule);

            foreach (var sch in collector)
            {
                // Skip internal schedules containing "<"
                if (sch.Name.Contains("<")) continue;

                string levelValue = "General";

                // Try to get a custom sorting parameter (optional)
                Parameter? param = sch.LookupParameter("FNX_TAB_TXT_Nivel 1");

                if (param != null && param.HasValue)
                {
                    string? val = param.AsString();
                    if (!string.IsNullOrWhiteSpace(val)) levelValue = val;
                }

                _allSchedules.Add(new ScheduleItem(sch, levelValue));
            }

            RefreshList();
        }

        private void RefreshList()
        {
            string search = SearchBox.Text?.ToLower() ?? "";

            // Filter list based on search text
            var filteredList = _allSchedules.Where(item =>
            {
                bool matchText = string.IsNullOrEmpty(search) || (item.ViewName != null && item.ViewName.ToLower().Contains(search));
                return matchText;
            }).ToList();

            ScheduleListBox.ItemsSource = filteredList;
        }

        // --- UI Events ---

        private void OnWindowDrag(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) DragMove();
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void OnSearchChanged(object sender, TextChangedEventArgs e)
        {
            RefreshList();
        }

        private void OnSelectAllClick(object sender, RoutedEventArgs e)
        {
            bool checkState = CheckSelectAll.IsChecked == true;
            if (ScheduleListBox.ItemsSource is List<ScheduleItem> currentItems)
            {
                foreach (var item in currentItems)
                {
                    item.IsSelected = checkState;
                }
            }
        }

        private void OnBrowseClick(object sender, RoutedEventArgs e)
        {
            string dateStr = DateTime.Now.ToString("yyyyMMdd");
            string fileNamePart = "Schedule";

            // Determine default filename based on first selection
            var firstSelected = _allSchedules.FirstOrDefault(x => x.IsSelected);
            if (firstSelected != null)
            {
                fileNamePart = CleanFileName(firstSelected.ViewName);
            }
            else
            {
                var firstVisible = ScheduleListBox.Items.Count > 0 ? ScheduleListBox.Items[0] as ScheduleItem : null;
                if (firstVisible != null) fileNamePart = CleanFileName(firstVisible.ViewName);
            }

            string defaultFileName = $"{dateStr}-{fileNamePart}.xlsx";

            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = defaultFileName,
                Title = "Select Destination File"
            };

            if (dialog.ShowDialog() == true)
            {
                _destinationPath = dialog.FileName;
                TxtPath.Text = _destinationPath;
                TxtPath.Foreground = System.Windows.Media.Brushes.Black;
            }
        }

        private string CleanFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "Schedule";
            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name;
        }

        private void OnExportClick(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_destinationPath))
            {
                Autodesk.Revit.UI.TaskDialog.Show("ZenBIM Warning", "Please select a destination folder first.");
                return;
            }

            var selectedItems = _allSchedules.Where(x => x.IsSelected).ToList();
            if (selectedItems.Count == 0)
            {
                Autodesk.Revit.UI.TaskDialog.Show("ZenBIM Warning", "No schedules selected.");
                return;
            }

            // UI Progress Display
            var progress = new ZenProgressView("ZenBIM: Export Excel", "Generating file...");
            progress.SetIcon("M4,2H20A2,2 0 0,1 22,4V20A2,2 0 0,1 20,22H4A2,2 0 0,1 2,20V4A2,2 0 0,1 4,2M4,4V8H20V4H4M4,10V20H20V10H4M8,12V18H16V12H8Z");
            progress.Show();

            try
            {
                this.Visibility = System.Windows.Visibility.Hidden;
                ExportToExcelClosedXML(selectedItems, progress);
                this.Close();
                Autodesk.Revit.UI.TaskDialog.Show("ZenBIM Success", "Export completed successfully!");
            }
            catch (Exception ex)
            {
                this.Visibility = System.Windows.Visibility.Visible;
                Autodesk.Revit.UI.TaskDialog.Show("ZenBIM Error", $"An error occurred:\n{ex.Message}");
            }
            finally
            {
                progress.Close();
            }
        }

        private void ExportToExcelClosedXML(List<ScheduleItem> items, ZenProgressView progress)
        {
            // Use ClosedXML to create workbook in memory (no Excel required)
            using (var workbook = new XLWorkbook())
            {
                int total = items.Count;
                int count = 0;

                foreach (var item in items)
                {
                    count++;
                    progress.Update(((double)count / total) * 100);

                    // 1. Sanitize Sheet Name
                    string cleanName = item.ViewName.Replace(":", "").Replace("/", "-").Replace("?", "").Replace("*", "").Replace("[", "").Replace("]", "");
                    if (cleanName.Length > 30) cleanName = cleanName.Substring(0, 30);

                    // Handle duplicate sheet names
                    string finalSheetName = cleanName;
                    int dup = 1;
                    while (workbook.Worksheets.Any(w => w.Name == finalSheetName))
                    {
                        string suffix = $" ({dup})";
                        int maxLen = 31 - suffix.Length;
                        if (cleanName.Length > maxLen) cleanName = cleanName.Substring(0, maxLen);
                        finalSheetName = cleanName + suffix;
                        dup++;
                    }

                    var worksheet = workbook.Worksheets.Add(finalSheetName);

                    // --- FONT DETECTION (Revit) ---
                    // Get specific font names used in Revit for Body and Header
                    string bodyFont = GetScheduleFont(_doc, item.RevitView, false) ?? "Calibri";
                    string headerFont = GetScheduleFont(_doc, item.RevitView, true) ?? "Calibri";

                    // --- GET REVIT TABLE DATA ---
                    TableData tableData = item.RevitView.GetTableData();
                    TableSectionData sectionHeader = tableData.GetSectionData(SectionType.Header);
                    TableSectionData sectionBody = tableData.GetSectionData(SectionType.Body);

                    // Start Position in Excel (B2)
                    int startRow = 2;
                    int startCol = 2;
                    int currentRow = startRow;

                    // ---------------------------------------------------------
                    // 2. EXPORT HEADERS (HEADER SECTION)
                    // ---------------------------------------------------------
                    if (sectionHeader != null)
                    {
                        int nHeaderRows = sectionHeader.NumberOfRows;
                        int nHeaderCols = sectionHeader.NumberOfColumns;

                        for (int r = 0; r < nHeaderRows; r++)
                        {
                            for (int c = 0; c < nHeaderCols; c++)
                            {
                                // Write Text
                                string cellText = item.RevitView.GetCellText(SectionType.Header, r, c);
                                var cell = worksheet.Cell(currentRow + r, startCol + c);
                                cell.Value = cellText;

                                // Apply Base Header Style
                                cell.Style.Font.FontName = headerFont;
                                cell.Style.Font.Bold = true;
                                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                cell.Style.Border.OutsideBorderColor = XLColor.Gray;
                                cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#F0F0F0"); // Light gray background

                                // --- MERGED CELLS LOGIC ---
                                // Check if this cell is the top-left of a merged region
                                TableMergedCell mergedCell = sectionHeader.GetMergedCell(r, c);
                                if (mergedCell.Left == c && mergedCell.Top == r)
                                {
                                    // If region is larger than 1x1
                                    if (mergedCell.Right != mergedCell.Left || mergedCell.Bottom != mergedCell.Top)
                                    {
                                        int spanRows = mergedCell.Bottom - mergedCell.Top + 1;
                                        int spanCols = mergedCell.Right - mergedCell.Left + 1;

                                        // Merge in Excel
                                        worksheet.Range(currentRow + r, startCol + c, currentRow + r + spanRows - 1, startCol + c + spanCols - 1).Merge();
                                    }
                                }

                                // Apply Revit specific visual overrides (colors, etc.)
                                TableCellStyle style = sectionHeader.GetTableCellStyle(r, c);
                                ApplyRevitStylesToCell(cell, style);
                            }
                        }
                        // Advance current row for Body section
                        currentRow += nHeaderRows;
                    }

                    // ---------------------------------------------------------
                    // 3. EXPORT DATA (BODY SECTION)
                    // ---------------------------------------------------------
                    if (sectionBody != null)
                    {
                        int nBodyRows = sectionBody.NumberOfRows;
                        int nBodyCols = sectionBody.NumberOfColumns;

                        for (int r = 0; r < nBodyRows; r++)
                        {
                            for (int c = 0; c < nBodyCols; c++)
                            {
                                string cellText = item.RevitView.GetCellText(SectionType.Body, r, c);
                                var cell = worksheet.Cell(currentRow + r, startCol + c);

                                // Data Type Detection (Number vs Text)
                                if (double.TryParse(cellText, out double numericVal))
                                    cell.Value = numericVal;
                                else
                                    cell.Value = cellText;

                                // Apply Global Body Font
                                cell.Style.Font.FontName = bodyFont;

                                // Apply Revit visual styles
                                TableCellStyle style = sectionBody.GetTableCellStyle(r, c);
                                ApplyRevitStylesToCell(cell, style);

                                // Default Borders
                                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                cell.Style.Border.OutsideBorderColor = XLColor.Gray;
                            }
                        }

                        // ---------------------------------------------------------
                        // 4. COLUMN WIDTHS
                        // ---------------------------------------------------------
                        for (int c = 0; c < nBodyCols; c++)
                        {
                            double widthInFeet = sectionBody.GetColumnWidth(c);
                            double width = widthInFeet * 120; // Approximation factor
                            if (width > 100) width = 100;
                            if (width < 8) width = 8;
                            worksheet.Column(startCol + c).Width = width;
                        }
                    }
                }

                // Save file
                workbook.SaveAs(_destinationPath);
            }

            // Logic to Open Excel automatically if CheckBox is checked
            if (CheckOpenExcel.IsChecked == true && File.Exists(_destinationPath))
            {
                try
                {
                    Process.Start(new ProcessStartInfo(_destinationPath) { UseShellExecute = true });
                }
                catch (Exception)
                {
                    // Ignore errors if system cannot open the file
                }
            }
        }

        /// <summary>
        /// Helper method to apply Revit TableCellStyle to ClosedXML Cell
        /// </summary>
        private void ApplyRevitStylesToCell(IXLCell cell, TableCellStyle style)
        {
            // Background Color
            if (style.BackgroundColor.IsValid)
            {
                // Ignore pure white (255,255,255) to avoid clutter
                if (!(style.BackgroundColor.Red == 255 && style.BackgroundColor.Green == 255 && style.BackgroundColor.Blue == 255))
                {
                    cell.Style.Fill.BackgroundColor = XLColor.FromArgb(
                        style.BackgroundColor.Red, style.BackgroundColor.Green, style.BackgroundColor.Blue);
                }
            }

            // Font Styles
            if (style.IsFontBold) cell.Style.Font.Bold = true;
            if (style.IsFontItalic) cell.Style.Font.Italic = true;
            if (style.IsFontUnderline) cell.Style.Font.Underline = XLFontUnderlineValues.Single;

            // Text Color
            if (style.TextColor.IsValid)
            {
                // Ignore pure black (0,0,0) as it is default
                if (!(style.TextColor.Red == 0 && style.TextColor.Green == 0 && style.TextColor.Blue == 0))
                {
                    cell.Style.Font.FontColor = XLColor.FromArgb(
                        style.TextColor.Red, style.TextColor.Green, style.TextColor.Blue);
                }
            }

            // Alignment
            switch (style.FontHorizontalAlignment)
            {
                case HorizontalAlignmentStyle.Left: cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left; break;
                case HorizontalAlignmentStyle.Center: cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; break;
                case HorizontalAlignmentStyle.Right: cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right; break;
            }
        }

        /// <summary>
        /// Retrieves the font name used in the schedule's appearance settings.
        /// Checks common parameter names in English and Spanish.
        /// </summary>
        private string? GetScheduleFont(Document doc, ViewSchedule vs, bool isHeader)
        {
            string[] paramNames = isHeader
                ? new[] { "Header Text", "Texto de encabezado", "Texto de título" }
                : new[] { "Body Text", "Texto de cuerpo" };

            foreach (var name in paramNames)
            {
                Parameter p = vs.LookupParameter(name);
                if (p != null && p.HasValue)
                {
                    ElementId typeId = p.AsElementId();
                    if (typeId != ElementId.InvalidElementId)
                    {
                        Element typeElem = doc.GetElement(typeId);
                        if (typeElem != null)
                        {
                            // The built-in parameter TEXT_FONT holds the font name (e.g., Arial)
                            Parameter fontParam = typeElem.get_Parameter(BuiltInParameter.TEXT_FONT);
                            if (fontParam != null && !string.IsNullOrEmpty(fontParam.AsString()))
                            {
                                return fontParam.AsString();
                            }
                        }
                    }
                }
            }
            return null;
        }
    }
}