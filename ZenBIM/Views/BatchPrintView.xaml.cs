//-----------------------------------------------------------------------------------------
// <copyright file="BatchPrintView.xaml.cs" company="plano y escala">
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
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;
using ZenBIM.Core;

namespace ZenBIM.Views
{
    public class SheetItem
    {
        public required string SheetNumber { get; set; }
        public required string SheetName { get; set; }
        public required ViewSheet Element { get; set; }
        public bool IsSelected { get; set; }
        public string PreviewName { get; set; } = "";
    }

    public class FilterItem
    {
        public required string Name { get; set; }
        public required ElementId Id { get; set; }
        public bool IsCollection { get; set; }
    }

    public class DWGSetupItem { public required string Name { get; set; } }

    public class NamingRulePreset
    {
        public string Rule { get; set; } = string.Empty;
        public string Separator { get; set; } = "-";
    }

    public partial class BatchPrintView : Window
    {
        private List<SheetItem> _allSheetItems = new List<SheetItem>();
        private Document _doc;
        private string _currentSeparator = "-";

        public List<ViewSheet> SelectedSheets => _allSheetItems.Where(s => s.IsSelected).Select(s => s.Element).ToList();
        public string NamingRule => TxtNamingRule.Text;
        public string OutputFolder { get; private set; } = string.Empty;

        public bool CombineFiles => CheckMerge.IsChecked == true;
        public bool OpenAfter => CheckOpen.IsChecked == true;
        public bool HideScopeBoxes => CheckLinks.IsChecked == true;
        public int ExportMode => ComboFormat.SelectedIndex;

        public string SelectedDWGSetup => (ComboDWGSetup.SelectedItem as DWGSetupItem)?.Name ?? "";
        public RasterQualityType SelectedQuality => ComboQuality.SelectedIndex == 2 ? RasterQualityType.High : RasterQualityType.Presentation;
        public ColorDepthType SelectedColor => ComboColors.SelectedIndex == 0 ? ColorDepthType.Color : ColorDepthType.GrayScale;

        public BatchPrintView(Document doc)
        {
            InitializeComponent();
            _doc = doc;

            LoadSettings();
            LoadDWGSetups();
            LoadFilters();
            LoadSheets();
        }

        private void LoadSettings()
        {
            var settings = BatchPrintSettings.Load();
            if (!string.IsNullOrEmpty(settings.LastNamingRule)) TxtNamingRule.Text = settings.LastNamingRule;
            else TxtNamingRule.Text = "{Sheet Number}-{Sheet Name}";

            if (!string.IsNullOrEmpty(settings.LastSeparator)) _currentSeparator = settings.LastSeparator;
            if (!string.IsNullOrEmpty(settings.LastOutputFolder) && Directory.Exists(settings.LastOutputFolder))
            {
                OutputFolder = settings.LastOutputFolder;
                TxtFolderPath.Text = OutputFolder;
            }
        }

        private void LoadFilters()
        {
            var filters = new List<FilterItem> { new FilterItem { Name = "Show All Sheets", Id = ElementId.InvalidElementId } };
            var sets = new FilteredElementCollector(_doc).OfClass(typeof(ViewSheetSet)).Cast<ViewSheetSet>().OrderBy(s => s.Name).ToList();
            foreach (var set in sets) filters.Add(new FilterItem { Name = $"Set: {set.Name}", Id = set.Id, IsCollection = false });

            try
            {
                var collections = new FilteredElementCollector(_doc).OfClass(typeof(SheetCollection)).Cast<SheetCollection>().OrderBy(s => s.Name).ToList();
                foreach (var col in collections) filters.Add(new FilterItem { Name = $"Collection: {col.Name}", Id = col.Id, IsCollection = true });
            }
            catch { }
            ComboFilter.ItemsSource = filters;
            ComboFilter.SelectedIndex = 0;
        }

        private void LoadDWGSetups()
        {
            var setups = new List<DWGSetupItem>();
            var elements = new FilteredElementCollector(_doc).OfClass(typeof(ExportDWGSettings)).Cast<ExportDWGSettings>().OrderBy(s => s.Name).ToList();
            foreach (var s in elements) setups.Add(new DWGSetupItem { Name = s.Name });
            if (setups.Count == 0) setups.Add(new DWGSetupItem { Name = "<In-Session>" });
            ComboDWGSetup.ItemsSource = setups;
            ComboDWGSetup.SelectedIndex = 0;
        }

        private void LoadSheets()
        {
            _allSheetItems.Clear();

            var collector = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(s => !s.IsPlaceholder);

            if (ComboFilter.SelectedItem is FilterItem selectedFilter && selectedFilter.Id != ElementId.InvalidElementId)
            {
                if (selectedFilter.IsCollection)
                    collector = collector.Where(s => s.SheetCollectionId == selectedFilter.Id);
                else
                {
                    var set = _doc.GetElement(selectedFilter.Id) as ViewSheetSet;
                    if (set != null && set.Views.Size > 0)
                    {
                        var viewIdsInSet = new HashSet<ElementId>();
                        foreach (Autodesk.Revit.DB.View v in set.Views) viewIdsInSet.Add(v.Id);
                        collector = collector.Where(s => viewIdsInSet.Contains(s.Id));
                    }
                }
            }

            var sorted = collector.OrderBy(s => s.SheetNumber).ToList();

            foreach (var s in sorted)
            {
                _allSheetItems.Add(new SheetItem
                {
                    SheetNumber = s.SheetNumber ?? "???",
                    SheetName = s.Name ?? "Unnamed",
                    Element = s,
                    IsSelected = false
                });
            }

            UpdatePreviews();
            RefreshList();
        }

        private void UpdatePreviews()
        {
            string rule = TxtNamingRule.Text;
            if (string.IsNullOrWhiteSpace(rule)) rule = "{Sheet Number}";

            foreach (var item in _allSheetItems)
            {
                string rawName = ParseTokens(rule, item.Element);
                foreach (char c in System.IO.Path.GetInvalidFileNameChars()) rawName = rawName.Replace(c, '_');

                string ext = ComboFormat.SelectedIndex == 0 ? ".pdf" : (ComboFormat.SelectedIndex == 1 ? ".dwg" : ".pdf/.dwg");
                item.PreviewName = rawName + ext;
            }
        }

        private string ParseTokens(string rule, ViewSheet sheet)
        {
            return Regex.Replace(rule, @"\{(.*?)\}", match =>
            {
                string paramName = match.Groups[1].Value;

                Parameter p = sheet.LookupParameter(paramName);
                if (p != null) return p.AsString() ?? p.AsValueString() ?? "";
                if (paramName == "Sheet Number") return sheet.SheetNumber;
                if (paramName == "Sheet Name") return sheet.Name;

                if (_doc.ProjectInformation != null)
                {
                    Parameter pInfo = _doc.ProjectInformation.LookupParameter(paramName);
                    if (pInfo != null) return pInfo.AsString() ?? pInfo.AsValueString() ?? "";
                    if (paramName == "Project Number") return _doc.ProjectInformation.Number;
                    if (paramName == "Project Name") return _doc.ProjectInformation.Name;
                }
                return "Unknown";
            });
        }

        private void RefreshList()
        {
            string search = TxtSearch.Text?.ToLower() ?? "";
            var filtered = _allSheetItems
                .Where(s => s.SheetNumber.ToLower().Contains(search) ||
                            s.SheetName.ToLower().Contains(search))
                .ToList();

            ListSheets.ItemsSource = filtered;
            ListSheets.Items.Refresh();
            UpdateCount();
        }

        private void UpdateCount() => TxtCount.Text = $"{_allSheetItems.Count(s => s.IsSelected)} sheets selected";

        private void OnFilterChanged(object sender, SelectionChangedEventArgs e) => LoadSheets();
        private void OnSearchChanged(object sender, TextChangedEventArgs e) => RefreshList();

        private void OnFormatChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PanelPDF == null || PanelDWG == null) return;
            int index = ComboFormat.SelectedIndex;

            // CORREGIDO: Uso explícito de System.Windows.Visibility para evitar CS0176
            PanelPDF.Visibility = (index == 0 || index == 2) ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            PanelDWG.Visibility = (index == 1 || index == 2) ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;

            if (CheckMerge != null) CheckMerge.IsEnabled = (index == 0 || index == 2);

            UpdatePreviews();
            if (ListSheets != null) ListSheets.Items.Refresh();
        }

        private void OnSelectAll(object sender, RoutedEventArgs e)
        {
            if (ListSheets.ItemsSource is List<SheetItem> visibleItems)
            {
                foreach (var item in visibleItems) item.IsSelected = true;
                ListSheets.Items.Refresh();
                UpdateCount();
            }
        }

        private void OnSelectNone(object sender, RoutedEventArgs e)
        {
            foreach (var item in _allSheetItems) item.IsSelected = false;
            ListSheets.Items.Refresh();
            UpdateCount();
        }

        private void OnAddParameterClick(object sender, RoutedEventArgs e)
        {
            var selector = new ParameterSelectorView(_doc, _currentSeparator, TxtNamingRule.Text);
            if (selector.ShowDialog() == true)
            {
                TxtNamingRule.Text = selector.ResultRule;
                _currentSeparator = selector.ResultSeparator;
                UpdatePreviews();
                ListSheets.Items.Refresh();
            }
        }

        private void OnRuleLostFocus(object sender, RoutedEventArgs e)
        {
            UpdatePreviews();
            ListSheets.Items.Refresh();
        }

        private void OnSaveRuleClick(object sender, RoutedEventArgs e)
        {
            var preset = new NamingRulePreset { Rule = TxtNamingRule.Text, Separator = _currentSeparator };
            var dialog = new Microsoft.Win32.SaveFileDialog { Title = "Save Preset", FileName = "NamingRule", DefaultExt = ".json", Filter = "JSON (*.json)|*.json" };
            if (dialog.ShowDialog() == true)
            {
                try
                {
                    File.WriteAllText(dialog.FileName, JsonSerializer.Serialize(preset, new JsonSerializerOptions { WriteIndented = true }));
                    // CORREGIDO: Uso explícito de System.Windows.MessageBox
                    System.Windows.MessageBox.Show("Preset saved!", "ZenBIM", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("Error: " + ex.Message, "ZenBIM", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void OnLoadRuleClick(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog { Title = "Load Preset", DefaultExt = ".json", Filter = "JSON (*.json)|*.json" };
            if (dialog.ShowDialog() == true)
            {
                try
                {
                    var preset = JsonSerializer.Deserialize<NamingRulePreset>(File.ReadAllText(dialog.FileName));
                    if (preset != null)
                    {
                        TxtNamingRule.Text = preset.Rule;
                        _currentSeparator = preset.Separator;
                        UpdatePreviews();
                        ListSheets.Items.Refresh();
                    }
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("Error: " + ex.Message, "ZenBIM", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void OnExportClick(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(OutputFolder))
            {
                System.Windows.MessageBox.Show("Select an output folder.", "ZenBIM", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (SelectedSheets.Count == 0)
            {
                System.Windows.MessageBox.Show("Select at least one sheet.", "ZenBIM", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var settings = new BatchPrintSettings
            {
                LastNamingRule = TxtNamingRule.Text,
                LastSeparator = _currentSeparator,
                LastOutputFolder = OutputFolder
            };
            BatchPrintSettings.Save(settings);

            this.DialogResult = true;
            this.Close();
        }

        private void OnBrowseFolder(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog { Description = "Select Output Folder", ShowNewFolderButton = true })
            {
                if (!string.IsNullOrEmpty(OutputFolder) && Directory.Exists(OutputFolder)) dialog.SelectedPath = OutputFolder;
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    OutputFolder = dialog.SelectedPath;
                    TxtFolderPath.Text = OutputFolder;
                }
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e) { if (e.ChangedButton == MouseButton.Left) this.DragMove(); }
        private void BtnClose_Click(object sender, RoutedEventArgs e) => this.Close();
    }
}