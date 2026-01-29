//-----------------------------------------------------------------------------------------
// <copyright file="ParameterSelectorView.xaml.cs" company="plano y escala">
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
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;

namespace ZenBIM.Views
{
    public partial class ParameterSelectorView : Window
    {
        public class ParamItem
        {
            public required string Name { get; set; }
            public required string Group { get; set; }
        }

        private List<ParamItem> _allParamsSource = new List<ParamItem>();

        public ObservableCollection<string> AvailableParams { get; set; } = new ObservableCollection<string>();
        public ObservableCollection<string> SelectedParams { get; set; } = new ObservableCollection<string>();

        public string ResultRule { get; private set; } = string.Empty;
        public string ResultSeparator { get; private set; } = "-";

        // Constructor modificado para recibir la regla actual
        public ParameterSelectorView(Document doc, string currentSeparator, string currentRule)
        {
            InitializeComponent();

            if (TxtSeparator != null) TxtSeparator.Text = currentSeparator;

            ListAvailable.ItemsSource = AvailableParams;
            ListSelected.ItemsSource = SelectedParams;

            LoadParameters(doc);
            PreloadCurrentRule(currentRule, currentSeparator);
        }

        private void LoadParameters(Document doc)
        {
            _allParamsSource.Clear();

            // 1. Sheet Parameters
            var sheet = new FilteredElementCollector(doc).OfClass(typeof(ViewSheet)).Cast<ViewSheet>().FirstOrDefault();
            if (sheet != null)
            {
                foreach (Parameter p in sheet.Parameters)
                {
                    if (!_allParamsSource.Any(x => x.Name == p.Definition.Name))
                        _allParamsSource.Add(new ParamItem { Name = p.Definition.Name, Group = "Sheet" });
                }
            }

            // Fallbacks
            if (!_allParamsSource.Any(x => x.Name == "Sheet Number")) _allParamsSource.Add(new ParamItem { Name = "Sheet Number", Group = "Sheet" });
            if (!_allParamsSource.Any(x => x.Name == "Sheet Name")) _allParamsSource.Add(new ParamItem { Name = "Sheet Name", Group = "Sheet" });

            // 2. Project Info Parameters
            if (doc.ProjectInformation != null)
            {
                foreach (Parameter p in doc.ProjectInformation.Parameters)
                {
                    if (!_allParamsSource.Any(x => x.Name == p.Definition.Name))
                        _allParamsSource.Add(new ParamItem { Name = p.Definition.Name, Group = "Project" });
                }
            }
            if (!_allParamsSource.Any(x => x.Name == "Project Number")) _allParamsSource.Add(new ParamItem { Name = "Project Number", Group = "Project" });

            _allParamsSource = _allParamsSource.OrderBy(x => x.Name).ToList();
            RefreshAvailableList();
        }

        private void PreloadCurrentRule(string rule, string separator)
        {
            if (string.IsNullOrWhiteSpace(rule)) return;

            try
            {
                // Simple parsing: split by separator, strip {}
                string[] tokens = rule.Split(new string[] { separator }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var t in tokens)
                {
                    string clean = t.Trim('{', '}');
                    // Add only if meaningful
                    if (!string.IsNullOrWhiteSpace(clean))
                    {
                        SelectedParams.Add(clean);
                    }
                }
                RefreshAvailableList();
            }
            catch { /* Ignore parsing errors */ }
        }

        private void RefreshAvailableList()
        {
            if (TxtSearch == null || ComboSource == null) return;

            AvailableParams.Clear();
            string query = TxtSearch.Text?.ToLower() ?? "";
            string filterMode = (ComboSource.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "All";

            foreach (var p in _allParamsSource)
            {
                if (!p.Name.ToLower().Contains(query)) continue;
                if (SelectedParams.Contains(p.Name)) continue;

                if (filterMode == "Sheet Parameters" && p.Group != "Sheet") continue;
                if (filterMode == "Project Info Parameters" && p.Group != "Project") continue;

                AvailableParams.Add(p.Name);
            }
        }

        // --- INTERACTION ---
        private void OnSearchChanged(object sender, TextChangedEventArgs e) => RefreshAvailableList();
        private void OnSourceChanged(object sender, SelectionChangedEventArgs e) => RefreshAvailableList();

        private void OnAddClick(object sender, RoutedEventArgs e)
        {
            var itemToAdd = ListAvailable.SelectedItem as string;
            if (!string.IsNullOrEmpty(itemToAdd))
            {
                SelectedParams.Add(itemToAdd);
                RefreshAvailableList();
            }
        }

        private void OnRemoveClick(object sender, RoutedEventArgs e)
        {
            var itemToRemove = ListSelected.SelectedItem as string;
            if (!string.IsNullOrEmpty(itemToRemove))
            {
                SelectedParams.Remove(itemToRemove);
                RefreshAvailableList();
            }
        }

        private void OnMoveUp(object sender, RoutedEventArgs e)
        {
            int index = ListSelected.SelectedIndex;
            if (index > 0) SelectedParams.Move(index, index - 1);
        }

        private void OnMoveDown(object sender, RoutedEventArgs e)
        {
            int index = ListSelected.SelectedIndex;
            if (index < SelectedParams.Count - 1 && index >= 0) SelectedParams.Move(index, index + 1);
        }

        private void OnApplyClick(object sender, RoutedEventArgs e)
        {
            if (SelectedParams.Count == 0)
            {
                System.Windows.MessageBox.Show("Please select at least one parameter.", "ZenBIM", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string sep = TxtSeparator.Text;
            var tokens = SelectedParams.Select(p => $"{{{p}}}");
            ResultRule = string.Join(sep, tokens);
            ResultSeparator = sep;

            this.DialogResult = true;
            this.Close();
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) this.DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e) => this.Close();
    }
}