//-----------------------------------------------------------------------------------------
// <copyright file="ColorSplasherView.xaml.cs" company="plano y escala">
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
using System.Windows.Media;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ZenBIM.Views
{
    public class ColorItem
    {
        public string ValueName { get; set; } = string.Empty;
        public int Count { get; set; } = 0;
        public SolidColorBrush ColorBrush { get; set; } = System.Windows.Media.Brushes.Gray;
        public Autodesk.Revit.DB.Color RevitColor { get; set; } = new Autodesk.Revit.DB.Color(128, 128, 128);
        public List<ElementId> ElementIds { get; set; } = new List<ElementId>();
    }

    public class ViewWrapper
    {
        public string Name { get; set; } = string.Empty;
        public Autodesk.Revit.DB.View RevitView { get; set; } = default!;
        public bool IsSelected { get; set; } = false;
    }

    public partial class ColorSplasherView : Window
    {
        private ExternalCommandData _commandData;
        private Document _doc;

        private List<Category> _allCategories = new List<Category>();
        private List<Parameter> _allParameters = new List<Parameter>();
        private List<ColorItem> _currentData = new List<ColorItem>();
        private Random _rnd = new Random();
        private List<ViewWrapper> _viewWrappers = new List<ViewWrapper>();

        public ColorSplasherView(ExternalCommandData commandData)
        {
            InitializeComponent();
            _commandData = commandData;
            _doc = commandData.Application.ActiveUIDocument.Document;
            LoadCategories();
        }

        // --- 1. DATA LOADING & FILTERS ---

        private void LoadCategories()
        {
            _allCategories = new List<Category>();
            Categories categories = _doc.Settings.Categories;

            foreach (Category c in categories)
            {
                if (c.CategoryType == CategoryType.Model && c.AllowsBoundParameters && c.IsVisibleInUI)
                {
                    _allCategories.Add(c);
                }
            }
            _allCategories = _allCategories.OrderBy(c => c.Name).ToList();
            RefreshCategoryList();
        }

        private void RefreshCategoryList()
        {
            string search = TxtSearchCat.Text.ToLower();
            ListCategories.ItemsSource = _allCategories.Where(c => c.Name.ToLower().Contains(search)).ToList();
            ListCategories.DisplayMemberPath = "Name";
        }

        private void OnCategorySearchChanged(object sender, TextChangedEventArgs e) => RefreshCategoryList();

        private void OnCategoryChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ListCategories.SelectedItem is Category selectedCat)
            {
                Element? firstElem = new FilteredElementCollector(_doc)
                    .OfCategoryId(selectedCat.Id)
                    .WhereElementIsNotElementType()
                    .FirstElement();

                if (firstElem != null)
                {
                    _allParameters = new List<Parameter>();
                    foreach (Parameter p in firstElem.Parameters)
                    {
                        if (p.StorageType == StorageType.String ||
                            p.StorageType == StorageType.Double ||
                            p.StorageType == StorageType.Integer ||
                            p.StorageType == StorageType.ElementId)
                        {
                            _allParameters.Add(p);
                        }
                    }
                    _allParameters = _allParameters.OrderBy(p => p.Definition.Name).ToList();

                    ListParameters.IsEnabled = true;
                    RefreshParameterList();
                }
                else
                {
                    ListParameters.ItemsSource = null;
                    ListParameters.IsEnabled = false;
                }
            }
        }

        private void RefreshParameterList()
        {
            string search = TxtSearchParam.Text.ToLower();
            ListParameters.ItemsSource = _allParameters.Where(p => p.Definition.Name.ToLower().Contains(search)).ToList();
            ListParameters.DisplayMemberPath = "Definition.Name";
        }

        private void OnParameterSearchChanged(object sender, TextChangedEventArgs e) => RefreshParameterList();
        private void OnParameterChanged(object sender, SelectionChangedEventArgs e) => CalculateData();

        private void OnShuffleClick(object sender, RoutedEventArgs e)
        {
            if (_currentData != null && _currentData.Count > 0)
            {
                foreach (var item in _currentData) AssignRandomColor(item);
                ListValues.ItemsSource = null;
                ListValues.ItemsSource = _currentData;
            }
        }

        // --- 2. LOGIC ---

        private void CalculateData()
        {
            if (ListCategories.SelectedItem is not Category cat ||
                ListParameters.SelectedItem is not Parameter paramSample) return;

            string paramName = paramSample.Definition.Name;
            var collector = new FilteredElementCollector(_doc).OfCategoryId(cat.Id).WhereElementIsNotElementType();
            Dictionary<string, List<ElementId>> groupedData = new Dictionary<string, List<ElementId>>();

            foreach (Element elem in collector)
            {
                Parameter? p = elem.LookupParameter(paramName);
                string val = "<Empty>";

                if (p != null && p.HasValue)
                {
                    val = p.AsValueString();
                    if (string.IsNullOrEmpty(val))
                    {
                        try { val = p.AsString(); } catch { }
                        if (string.IsNullOrEmpty(val)) val = "<Empty>";
                    }
                }
                if (!groupedData.ContainsKey(val)) groupedData[val] = new List<ElementId>();
                groupedData[val].Add(elem.Id);
            }

            _currentData = new List<ColorItem>();
            foreach (var kvp in groupedData)
            {
                var item = new ColorItem { ValueName = kvp.Key, Count = kvp.Value.Count, ElementIds = kvp.Value };
                AssignRandomColor(item);
                _currentData.Add(item);
            }
            _currentData = _currentData.OrderBy(x => x.ValueName).ToList();
            ListValues.ItemsSource = _currentData;

            // Updated status text in English
            if (TxtStatus != null)
                TxtStatus.Text = $"{_currentData.Count} unique values found.";
        }

        private void AssignRandomColor(ColorItem item)
        {
            // Vibrant colors (Light Theme compatible)
            byte r = (byte)_rnd.Next(40, 240);
            byte g = (byte)_rnd.Next(40, 240);
            byte b = (byte)_rnd.Next(40, 240);
            item.ColorBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(r, g, b));
            item.RevitColor = new Autodesk.Revit.DB.Color(r, g, b);
        }

        // --- 3. APPLY ---

        private void ApplyToViews(List<Autodesk.Revit.DB.View> views)
        {
            if (_currentData == null || _currentData.Count == 0) return;

            FillPatternElement? solidPattern = new FilteredElementCollector(_doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(p => p.GetFillPattern().IsSolidFill);

            if (solidPattern == null) return;

            using (Transaction t = new Transaction(_doc, "ZenBIM: Color Splasher"))
            {
                t.Start();
                try
                {
                    foreach (Autodesk.Revit.DB.View view in views)
                    {
                        foreach (var item in _currentData)
                        {
                            OverrideGraphicSettings ogs = new OverrideGraphicSettings();
                            ogs.SetSurfaceForegroundPatternId(solidPattern.Id);
                            ogs.SetSurfaceForegroundPatternColor(item.RevitColor);
                            ogs.SetProjectionLineColor(item.RevitColor);

                            if (view.ViewType != ViewType.ThreeD)
                            {
                                ogs.SetCutForegroundPatternId(solidPattern.Id);
                                ogs.SetCutForegroundPatternColor(item.RevitColor);
                                ogs.SetCutLineColor(item.RevitColor);
                            }

                            foreach (ElementId eid in item.ElementIds)
                            {
                                try { view.SetElementOverrides(eid, ogs); } catch { }
                            }
                        }
                    }
                    t.Commit();

                    OverlayViewSelector.Visibility = System.Windows.Visibility.Collapsed;
                    OverlaySuccess.Visibility = System.Windows.Visibility.Visible;
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    System.Windows.MessageBox.Show(ex.Message);
                }
            }
        }

        private void OnApplyCurrentClick(object sender, RoutedEventArgs e)
        {
            ApplyToViews(new List<Autodesk.Revit.DB.View> { _doc.ActiveView });
        }

        // --- 4. OVERLAYS MANAGEMENT ---

        private void OpenViewSelector(object sender, RoutedEventArgs e)
        {
            var views = new FilteredElementCollector(_doc)
                .OfClass(typeof(Autodesk.Revit.DB.View))
                .Cast<Autodesk.Revit.DB.View>()
                .Where(v => !v.IsTemplate && v.ViewType != ViewType.Legend && v.ViewType != ViewType.Schedule && v.ViewType != ViewType.ProjectBrowser)
                .OrderBy(v => v.Name)
                .ToList();

            _viewWrappers = views.Select(v => new ViewWrapper { Name = v.Name, RevitView = v, IsSelected = false }).ToList();

            TxtSearchView.Text = "";
            ListViews.ItemsSource = _viewWrappers;

            OverlayViewSelector.Visibility = System.Windows.Visibility.Visible;
        }

        private void OnViewSearchChanged(object sender, TextChangedEventArgs e)
        {
            string filter = TxtSearchView.Text.ToLower();
            ListViews.ItemsSource = _viewWrappers.Where(v => v.Name.ToLower().Contains(filter)).ToList();
        }

        private void ApplySelectedViews(object sender, RoutedEventArgs e)
        {
            var selected = _viewWrappers.Where(v => v.IsSelected).Select(v => v.RevitView).ToList();
            if (selected.Count > 0)
            {
                ApplyToViews(selected);
            }
        }

        private void CloseOverlay(object sender, RoutedEventArgs e)
        {
            OverlayViewSelector.Visibility = System.Windows.Visibility.Collapsed;
            OverlaySuccess.Visibility = System.Windows.Visibility.Collapsed;
        }

        // --- 5. LEGEND & RESET ---

        private void OnCreateLegendClick(object sender, RoutedEventArgs e)
        {
            if (_currentData == null || _currentData.Count == 0) return;

            FilledRegionType? regionType = new FilteredElementCollector(_doc).OfClass(typeof(FilledRegionType)).Cast<FilledRegionType>().FirstOrDefault();
            if (regionType == null) return;

            double width = UnitUtils.ConvertToInternalUnits(150, UnitTypeId.Centimeters);
            double height = UnitUtils.ConvertToInternalUnits(75, UnitTypeId.Centimeters);
            double spacing = height * 1.5;
            double textOffset = UnitUtils.ConvertToInternalUnits(18, UnitTypeId.Centimeters);

            TextNoteType? textType = new FilteredElementCollector(_doc).OfClass(typeof(TextNoteType)).Cast<TextNoteType>().FirstOrDefault();

            using (Transaction t = new Transaction(_doc, "ZenBIM: Create Legend"))
            {
                t.Start();
                try
                {
                    var existingLegend = new FilteredElementCollector(_doc)
                        .OfClass(typeof(Autodesk.Revit.DB.View))
                        .Cast<Autodesk.Revit.DB.View>()
                        .FirstOrDefault(v => v.ViewType == ViewType.Legend && !v.IsTemplate);

                    if (existingLegend == null) { System.Windows.MessageBox.Show("No Legend View found to duplicate."); return; }

                    Autodesk.Revit.DB.View legendView = (Autodesk.Revit.DB.View)_doc.GetElement(existingLegend.Duplicate(ViewDuplicateOption.Duplicate));

                    if (legendView != null)
                    {
                        try { legendView.Name = $"Color Legend {DateTime.Now:HHmm}"; } catch { }
                        legendView.Scale = 100;

                        XYZ cursor = XYZ.Zero;
                        FillPatternElement? solid = new FilteredElementCollector(_doc).OfClass(typeof(FillPatternElement)).Cast<FillPatternElement>().FirstOrDefault(x => x.GetFillPattern().IsSolidFill);

                        if (solid != null)
                        {
                            foreach (var item in _currentData)
                            {
                                List<CurveLoop> loops = new List<CurveLoop>();
                                CurveLoop loop = new CurveLoop();
                                loop.Append(Line.CreateBound(cursor, new XYZ(cursor.X + width, cursor.Y, 0)));
                                loop.Append(Line.CreateBound(new XYZ(cursor.X + width, cursor.Y, 0), new XYZ(cursor.X + width, cursor.Y - height, 0)));
                                loop.Append(Line.CreateBound(new XYZ(cursor.X + width, cursor.Y - height, 0), new XYZ(cursor.X, cursor.Y - height, 0)));
                                loop.Append(Line.CreateBound(new XYZ(cursor.X, cursor.Y - height, 0), cursor));
                                loops.Add(loop);

                                FilledRegion region = FilledRegion.Create(_doc, regionType.Id, legendView.Id, loops);

                                OverrideGraphicSettings ogs = new OverrideGraphicSettings();
                                ogs.SetSurfaceForegroundPatternId(solid.Id);
                                ogs.SetSurfaceForegroundPatternColor(item.RevitColor);
                                ogs.SetProjectionLineColor(new Autodesk.Revit.DB.Color(255, 255, 255));
                                legendView.SetElementOverrides(region.Id, ogs);

                                if (textType != null)
                                {
                                    TextNote.Create(_doc, legendView.Id, new XYZ(cursor.X + width + (width * 0.2), cursor.Y - (height / 2) + textOffset, 0), $"{item.ValueName} ({item.Count})", textType.Id);
                                }
                                cursor = new XYZ(cursor.X, cursor.Y - spacing, 0);
                            }
                        }
                        t.Commit();
                        _commandData.Application.ActiveUIDocument.ActiveView = legendView;
                        this.Close();
                    }
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    System.Windows.MessageBox.Show(ex.Message);
                }
            }
        }

        private void OnResetClick(object sender, RoutedEventArgs e)
        {
            using (Transaction t = new Transaction(_doc, "ZenBIM: Reset Colors"))
            {
                t.Start();
                if (_currentData != null)
                {
                    foreach (var item in _currentData)
                    {
                        foreach (var eid in item.ElementIds)
                        {
                            try { _doc.ActiveView.SetElementOverrides(eid, new OverrideGraphicSettings()); } catch { }
                        }
                    }
                }
                t.Commit();
            }
            _commandData.Application.ActiveUIDocument.RefreshActiveView();
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e) { if (e.ChangedButton == MouseButton.Left) DragMove(); }
        private void BtnClose_Click(object sender, RoutedEventArgs e) => Close();
    }
}