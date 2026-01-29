//-----------------------------------------------------------------------------------------
// <copyright file="ImportExcelView.xaml.cs" company="plano y escala">
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
using System.Windows;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace ZenBIM.Views
{
    public partial class ImportExcelView : Window
    {
        public string SelectedFilePath { get; private set; } = string.Empty;
        public string SelectedSheetName { get; private set; } = string.Empty;
        public string SelectedRangeName { get; private set; } = string.Empty;

        // Diccionario para guardar la relación: Nombre Hoja -> Lista de Rangos
        private Dictionary<string, List<string>> _sheetMap = new Dictionary<string, List<string>>();

        public ImportExcelView()
        {
            InitializeComponent();
        }

        private void OnWindowDrag(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left) DragMove();
        }

        private void OnCloseClick(object sender, RoutedEventArgs e) { Close(); }

        private void OnBrowseClick(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog { Filter = "Excel Files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm" };
            if (dialog.ShowDialog() == true)
            {
                string path = dialog.FileName;

                // 1. Check if file is locked BEFORE processing
                if (CheckFileLock(path))
                {
                    return; // Stop if user cancelled or couldn't close file
                }

                SelectedFilePath = path;
                TxtPath.Text = Path.GetFileName(SelectedFilePath);
                TxtPath.Foreground = System.Windows.Media.Brushes.Black;

                LoadExcelMetadata(SelectedFilePath);
            }
        }

        private bool CheckFileLock(string path)
        {
            while (IsFileLocked(path))
            {
                // CORRECCIÓN: Uso explícito de System.Windows.MessageBox y manejo del resultado
                System.Windows.MessageBoxResult result = System.Windows.MessageBox.Show(
                    "The selected Excel file is open.\n\nPlease close it to continue.",
                    "ZenBIM: File in Use",
                    System.Windows.MessageBoxButton.OKCancel,
                    System.Windows.MessageBoxImage.Warning);

                if (result == System.Windows.MessageBoxResult.Cancel) return true; // User gave up
            }
            return false; // File is free
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

        private void LoadExcelMetadata(string path)
        {
            try
            {
                System.Windows.Input.Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
                _sheetMap.Clear();

                using (var workbook = new XLWorkbook(path))
                {
                    foreach (var sheet in workbook.Worksheets)
                    {
                        string sName = sheet.Name;
                        if (!_sheetMap.ContainsKey(sName))
                            _sheetMap[sName] = new List<string>();

                        var relevantRanges = new List<string>();

                        // 1. Global Defined Names pointing to this sheet
                        foreach (var dn in workbook.DefinedNames)
                        {
                            if (dn.Ranges.Any())
                            {
                                var rng = dn.Ranges.FirstOrDefault();
                                if (rng != null && rng.Worksheet.Name == sName)
                                {
                                    relevantRanges.Add(dn.Name);
                                }
                            }
                        }

                        // 2. Sheet scoped Named Ranges
                        // CORRECCIÓN: Usar DefinedNames en lugar de NamedRanges (Obsoleto)
                        foreach (var dn in sheet.DefinedNames)
                        {
                            relevantRanges.Add(dn.Name);
                        }

                        _sheetMap[sName] = relevantRanges.Distinct().OrderBy(x => x).ToList();
                    }
                }

                // Populate Sheets Combo
                ComboSheets.ItemsSource = _sheetMap.Keys.OrderBy(k => k).ToList();
                ComboSheets.IsEnabled = _sheetMap.Count > 0;

                if (ComboSheets.Items.Count > 0)
                    ComboSheets.SelectedIndex = 0;
                else
                    ComboRanges.ItemsSource = null;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Error reading file metadata: " + ex.Message);
            }
            finally
            {
                System.Windows.Input.Mouse.OverrideCursor = null;
            }
        }

        private void OnSheetChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            ComboRanges.ItemsSource = null;
            ComboRanges.IsEnabled = false;

            if (ComboSheets.SelectedItem == null) return;

            string? sheetName = ComboSheets.SelectedItem.ToString();

            // CORRECCIÓN: Validación estricta de nulos para el diccionario
            if (!string.IsNullOrEmpty(sheetName) && _sheetMap.TryGetValue(sheetName, out List<string>? ranges))
            {
                ComboRanges.ItemsSource = ranges;
                ComboRanges.IsEnabled = ranges != null && ranges.Count > 0;
                if (ranges != null && ranges.Count > 0) ComboRanges.SelectedIndex = 0;
            }
        }

        private void OnImportClick(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(SelectedFilePath))
            {
                System.Windows.MessageBox.Show("Please select a file.");
                return;
            }
            if (ComboSheets.SelectedItem == null)
            {
                System.Windows.MessageBox.Show("Please select a worksheet.");
                return;
            }
            if (ComboRanges.SelectedItem == null)
            {
                System.Windows.MessageBox.Show("Please select a Named Range to import.");
                return;
            }

            if (CheckFileLock(SelectedFilePath)) return;

            // CORRECCIÓN: Asignación segura
            SelectedSheetName = ComboSheets.SelectedItem?.ToString() ?? string.Empty;
            SelectedRangeName = ComboRanges.SelectedItem?.ToString() ?? string.Empty;

            this.DialogResult = true;
            this.Close();
        }
    }
}