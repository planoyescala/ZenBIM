//-----------------------------------------------------------------------------------------
// <copyright file="ReorderingView.xaml.cs" company="plano y escala">
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
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
// --- FIX: Importamos el espacio de nombres donde está el Handler ---
using ZenBIM.Commands;

namespace ZenBIM.Views
{
    public class ReorderItem
    {
        public Element? Element { get; set; }
        public string ElementId { get; set; } = string.Empty;
        public string CurrentValue { get; set; } = string.Empty;
        public string NewValue { get; set; } = string.Empty;
        public XYZ LocationPoint { get; set; } = XYZ.Zero;
    }

    public partial class ReorderingView : Window
    {
        private ExternalCommandData _commandData;
        private Document _doc;

        // External Event Variables
        private ManualRenumberHandler _manualHandler;
        private ExternalEvent _manualEvent;

        private List<Category> _allCategories = new List<Category>();
        private List<ReorderItem> _previewItems = new List<ReorderItem>();
        private bool _isRoomCategorySelected = false;

        public ReorderingView(ExternalCommandData commandData)
        {
            InitializeComponent();
            _commandData = commandData;
            _doc = commandData.Application.ActiveUIDocument.Document;

            // Initialize Handler and External Event
            _manualHandler = new ManualRenumberHandler();
            _manualEvent = ExternalEvent.Create(_manualHandler);

            LoadCategories();
        }

        private void LoadCategories()
        {
            _allCategories = new List<Category>();
            Categories categories = _doc.Settings.Categories;
            foreach (Category c in categories)
            {
                if (c.CategoryType == CategoryType.Model && c.AllowsBoundParameters && c.IsVisibleInUI)
                    _allCategories.Add(c);
            }
            _allCategories = _allCategories.OrderBy(c => c.Name).ToList();
            CmbCategory.ItemsSource = _allCategories;
        }

        private void CmbCategory_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CmbCategory.SelectedItem is not Category selectedCat) return;
            _isRoomCategorySelected = (selectedCat.Id.Value == (long)BuiltInCategory.OST_Rooms);

            if (_isRoomCategorySelected)
            {
                PanelRoomFilters.Visibility = System.Windows.Visibility.Visible;
                LoadRoomFilterParameters(selectedCat);
            }
            else
            {
                PanelRoomFilters.Visibility = System.Windows.Visibility.Collapsed;
            }
            LoadTargetParameters(selectedCat);
            RecalculatePreview(null, null);
        }

        private void LoadRoomFilterParameters(Category cat)
        {
            Element? firstElem = new FilteredElementCollector(_doc).OfCategoryId(cat.Id).WhereElementIsNotElementType().FirstElement();
            if (firstElem == null) return;
            var filterParams = new List<Parameter>();
            foreach (Parameter p in firstElem.Parameters)
            {
                if (p.StorageType == StorageType.String || p.StorageType == StorageType.ElementId) filterParams.Add(p);
            }
            CmbRoomFilterParam.ItemsSource = filterParams.OrderBy(p => p.Definition.Name).ToList();
            CmbRoomFilterParam.DisplayMemberPath = "Definition.Name";
        }

        private void LoadTargetParameters(Category cat)
        {
            Element? firstElem = new FilteredElementCollector(_doc).OfCategoryId(cat.Id).WhereElementIsNotElementType().FirstElement();
            if (firstElem == null) { CmbTargetParam.ItemsSource = null; return; }
            var writeableParams = new List<Parameter>();
            foreach (Parameter p in firstElem.Parameters)
            {
                if (!p.IsReadOnly && (p.StorageType == StorageType.String || p.StorageType == StorageType.Integer)) writeableParams.Add(p);
            }
            CmbTargetParam.ItemsSource = writeableParams.OrderBy(p => p.Definition.Name).ToList();
            var defaultParam = writeableParams.FirstOrDefault(p => p.Definition.Name == "Mark" || p.Definition.Name == "Marca" || p.Definition.Name == "Number" || p.Definition.Name == "Número");
            if (defaultParam != null) CmbTargetParam.SelectedItem = defaultParam;
            else if (writeableParams.Count > 0) CmbTargetParam.SelectedIndex = 0;
        }

        private void CmbRoomFilterParam_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CmbRoomFilterParam.SelectedItem is not Parameter selectedParam) return;
            var col = new FilteredElementCollector(_doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType();
            HashSet<string> values = new HashSet<string>();
            foreach (Element elem in col) values.Add(GetParamValue(elem, selectedParam.Definition.Name));
            CmbRoomFilterValue.ItemsSource = values.OrderBy(x => x).ToList();
            if (values.Count > 0) CmbRoomFilterValue.SelectedIndex = 0;
        }
        private void CmbRoomFilterValue_SelectionChanged(object sender, SelectionChangedEventArgs e) => RecalculatePreview(null, null);

        // --- CALCULATION LOGIC ---
        private void RecalculatePreview(object? sender, RoutedEventArgs? e)
        {
            if (CmbCategory.SelectedItem is not Category selectedCat) return;
            if (CmbTargetParam.SelectedItem is not Parameter targetParamDef) return;

            var collector = new FilteredElementCollector(_doc).OfCategoryId(selectedCat.Id).WhereElementIsNotElementType().ToElements();
            var rawList = new List<ReorderItem>();

            string? filterParamName = (CmbRoomFilterParam.SelectedItem as Parameter)?.Definition.Name;
            string? filterValue = CmbRoomFilterValue.SelectedItem as string;

            foreach (Element elem in collector)
            {
                if (_isRoomCategorySelected && !string.IsNullOrEmpty(filterParamName) && filterValue != null)
                {
                    if (GetParamValue(elem, filterParamName) != filterValue) continue;
                }

                XYZ loc = XYZ.Zero;
                if (elem.Location is LocationPoint lp) loc = lp.Point;
                else if (elem.Location is LocationCurve lc) loc = (lc.Curve.GetEndPoint(0) + lc.Curve.GetEndPoint(1)) / 2;
                else if (elem.get_BoundingBox(null) != null) { var bb = elem.get_BoundingBox(null); loc = (bb.Min + bb.Max) / 2; }

                rawList.Add(new ReorderItem { Element = elem, ElementId = elem.Id.ToString(), CurrentValue = GetParamValue(elem, targetParamDef.Definition.Name), LocationPoint = loc });
            }

            int sortIndex = CmbSortMethod.SelectedIndex;
            if (sortIndex == 0) rawList = rawList.OrderBy(x => x.LocationPoint.X).ThenBy(x => x.LocationPoint.Y).ToList();
            else if (sortIndex == 1) rawList = rawList.OrderBy(x => x.LocationPoint.Y).ThenBy(x => x.LocationPoint.X).ToList();
            else rawList = rawList.OrderBy(x => x.Element!.Id.Value).ToList();

            // Padding logic
            string prefix = TxtPrefix.Text ?? "";
            string suffix = TxtSuffix.Text ?? "";
            int.TryParse(TxtStart.Text, out int start);
            int.TryParse(TxtStep.Text, out int step);

            int padLength = TxtStart.Text.Length;
            bool shouldPad = TxtStart.Text.StartsWith("0") && TxtStart.Text.Length > 1;

            int currentNum = start;
            foreach (var item in rawList)
            {
                string numStr = currentNum.ToString();
                if (shouldPad) numStr = numStr.PadLeft(padLength, '0');
                item.NewValue = $"{prefix}{numStr}{suffix}";
                currentNum += step;
            }

            _previewItems = rawList;
            GridPreview.ItemsSource = _previewItems;
            TxtCount.Text = $"{_previewItems.Count} elements found";
        }

        // ========================================================
        // MANUAL MODE BUTTON
        // ========================================================
        private void OnManualModeClick(object sender, RoutedEventArgs e)
        {
            if (CmbCategory.SelectedItem is not Category cat || CmbTargetParam.SelectedItem is not Parameter param)
            {
                System.Windows.MessageBox.Show("Select Category and Parameter first.");
                return;
            }

            // 1. Pass data to handler
            _manualHandler.MainWindow = this;
            _manualHandler.TargetCategory = cat;
            _manualHandler.TargetParameter = param;
            _manualHandler.Prefix = TxtPrefix.Text ?? "";
            _manualHandler.Suffix = TxtSuffix.Text ?? "";

            int.TryParse(TxtStart.Text, out int start);
            _manualHandler.StartNumber = start;

            int.TryParse(TxtStep.Text, out int step);
            _manualHandler.Step = step;

            _manualHandler.PadLength = TxtStart.Text.Length;
            _manualHandler.PadZeros = TxtStart.Text.StartsWith("0") && TxtStart.Text.Length > 1;

            // 2. Raise External Event
            _manualEvent.Raise();
        }

        // ========================================================
        // AUTO APPLY BUTTON
        // ========================================================
        private void OnApplyClick(object sender, RoutedEventArgs e)
        {
            if (_previewItems == null || _previewItems.Count == 0) return;
            if (CmbTargetParam.SelectedItem is not Parameter targetParamSample) return;

            string paramName = targetParamSample.Definition.Name;

            try
            {
                using (Transaction t = new Transaction(_doc, "ZenBIM: Auto Reorder"))
                {
                    t.Start();
                    foreach (var item in _previewItems)
                    {
                        try
                        {
                            Parameter? p = item.Element?.LookupParameter(paramName);
                            if (p != null && !p.IsReadOnly)
                            {
                                if (p.StorageType == StorageType.String) p.Set(item.NewValue);
                                else if (p.StorageType == StorageType.Integer && int.TryParse(item.NewValue, out int iVal)) p.Set(iVal);
                            }
                        }
                        catch { }
                    }
                    t.Commit();
                    System.Windows.MessageBox.Show("Automatic renumbering completed.");
                    RecalculatePreview(null, null);
                }
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException)
            {
                System.Windows.MessageBox.Show("Context Error: To apply automatic changes in this modeless window, we need to convert this to an External Event as well. (Pending implementation).");
            }
        }

        private string GetParamValue(Element elem, string paramName)
        {
            Parameter? p = elem.LookupParameter(paramName);
            if (p == null) return "";
            if (p.StorageType == StorageType.String) return p.AsString() ?? "";
            return p.AsValueString() ?? "";
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e) { if (e.ChangedButton == MouseButton.Left) DragMove(); }
        private void BtnClose_Click(object sender, RoutedEventArgs e) => Close();
    }
}