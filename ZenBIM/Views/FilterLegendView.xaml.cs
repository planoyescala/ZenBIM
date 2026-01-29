//-----------------------------------------------------------------------------------------
// <copyright file="FilterLegendView.xaml.cs" company="plano y escala">
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

namespace ZenBIM.Views
{
    // Simple wrapper to handle View Object vs Display logic
    public class LegendViewWrapper
    {
        public string Name { get; set; } = string.Empty;
        public Autodesk.Revit.DB.View View { get; set; } = default!;
        public bool IsTemplate { get; set; }
    }

    public partial class FilterLegendView : Window
    {
        private List<LegendViewWrapper> _allWrappers = new List<LegendViewWrapper>();

        // Robust property to retrieve selected view
        public Autodesk.Revit.DB.View? SelectedView =>
            (FilterListBox.SelectedItem as LegendViewWrapper)?.View;

        public TextNoteType? SelectedTextType => ComboTextTypes.SelectedItem as TextNoteType;

        // Numeric properties connected to Inputs
        public double RegionWidth => double.TryParse(TxtRegionWidth.Text, out double d) ? d : 150;
        public double RegionHeight => double.TryParse(TxtRegionHeight.Text, out double d) ? d : 75;
        public double LineWidth => double.TryParse(TxtLineWidth.Text, out double d) ? d : 150;

        public FilterLegendView(Document doc)
        {
            InitializeComponent();
            LoadData(doc);
        }

        private void LoadData(Document doc)
        {
            // 1. Text Types
            var textTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .OrderBy(x => x.Name)
                .ToList();

            ComboTextTypes.ItemsSource = textTypes;
            ComboTextTypes.DisplayMemberPath = "Name"; // ComboBox still needs this or ItemTemplate

            // Smart Selection (Century Gothic or First available)
            var defaultType = textTypes.FirstOrDefault(t => t.Name.Contains("Century Gothic") || t.Name.Contains("Arial"));
            if (defaultType != null) ComboTextTypes.SelectedItem = defaultType;
            else if (textTypes.Count > 0) ComboTextTypes.SelectedIndex = 0;

            // 2. Views (Valid for filters only)
            var views = new FilteredElementCollector(doc)
                .OfClass(typeof(Autodesk.Revit.DB.View))
                .Cast<Autodesk.Revit.DB.View>()
                .Where(v => v.ViewType != ViewType.Internal &&
                            v.ViewType != ViewType.ProjectBrowser &&
                            v.ViewType != ViewType.Legend &&
                            v.ViewType != ViewType.DrawingSheet &&
                            v.ViewType != ViewType.Schedule)
                .ToList();

            _allWrappers = views.Select(v => new LegendViewWrapper
            {
                Name = v.Name,
                View = v,
                IsTemplate = v.IsTemplate
            }).OrderBy(x => x.Name).ToList();

            RefreshList();
        }

        private void OnSearchChanged(object sender, TextChangedEventArgs e) => RefreshList();
        private void OnFilterChanged(object sender, RoutedEventArgs e) => RefreshList();

        private void RefreshList()
        {
            if (FilterListBox == null) return;

            bool showViews = CheckViews?.IsChecked == true;
            bool showTemplates = CheckTemplates?.IsChecked == true;
            string searchText = SearchBox.Text?.ToLower() ?? "";

            var filtered = _allWrappers
                .Where(w => (w.IsTemplate && showTemplates) || (!w.IsTemplate && showViews))
                .Where(w => w.Name.ToLower().Contains(searchText))
                .ToList();

            // Note: DisplayMemberPath removed here because we use ItemTemplate in XAML for better styling
            FilterListBox.ItemsSource = filtered;
        }

        private void OnCreateClick(object sender, RoutedEventArgs e)
        {
            if (SelectedView == null)
            {
                // Native MessageBox used for simplicity
                System.Windows.MessageBox.Show("Please, select a source view.", "ZenBIM", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Validate numbers
            if (RegionWidth <= 0 || RegionHeight <= 0 || LineWidth <= 0)
            {
                System.Windows.MessageBox.Show("Dimensions must be greater than 0.", "ZenBIM", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            this.DialogResult = true;
            this.Close();
        }

        // --- Window Events ---
        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) this.DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e) => this.Close();
    }
}