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
using System.Windows.Controls; // Main Namespace for WPF controls
using System.Windows.Input;
using Autodesk.Revit.DB;

namespace ZenBIM.Views
{
    public partial class ParameterSelectorView : Window
    {
        // Class compatible with all Revit versions
        public class ParamItem
        {
            public string Name { get; set; } = string.Empty;
            public string Group { get; set; } = string.Empty;
        }

        private List<ParamItem> _allParamsSource = new List<ParamItem>();

        public ObservableCollection<string> AvailableParams { get; set; } = new ObservableCollection<string>();
        public ObservableCollection<string> SelectedParams { get; set; } = new ObservableCollection<string>();

        public string ResultRule { get; private set; } = string.Empty;
        public string ResultSeparator { get; private set; } = "-";

        public ParameterSelectorView(Document doc, string currentSeparator, string currentRule)
        {
            InitializeComponent();

            // --- FIX START: Connect Data Collections to UI ListBoxes ---
            ListAvailable.ItemsSource = AvailableParams;
            ListSelected.ItemsSource = SelectedParams;
            // --- FIX END ---

            LoadParameters(doc);

            if (!string.IsNullOrEmpty(currentSeparator))
            {
                TxtSeparator.Text = currentSeparator;
            }

            ParseCurrentRule(currentRule);
        }

        private void LoadParameters(Document doc)
        {
            _allParamsSource.Clear();

            // 1. Built-in Sheet Parameters
            AddParam("Sheet Number", "Built-in");
            AddParam("Sheet Name", "Built-in");

            // 2. Project Info Parameters
            AddParam("Project Number", "Project Info");
            AddParam("Project Name", "Project Info");

            // 3. Sheet Parameters (Instance)
            var sheetCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .FirstElement();

            if (sheetCollector != null)
            {
                foreach (Parameter p in sheetCollector.Parameters)
                {
                    if (p.Definition != null)
                    {
                        AddParam(p.Definition.Name, "Sheet Parameter");
                    }
                }
            }

            // 4. Project Information Parameters
            if (doc.ProjectInformation != null)
            {
                foreach (Parameter p in doc.ProjectInformation.Parameters)
                {
                    if (p.Definition != null)
                    {
                        AddParam(p.Definition.Name, "Project Information");
                    }
                }
            }

            // Sort and Remove Duplicates
            _allParamsSource = _allParamsSource
                .GroupBy(x => x.Name)
                .Select(g => g.First())
                .OrderBy(x => x.Name)
                .ToList();
        }

        private void AddParam(string name, string group)
        {
            _allParamsSource.Add(new ParamItem { Name = name, Group = group });
        }

        private void ParseCurrentRule(string rule)
        {
            SelectedParams.Clear();
            if (string.IsNullOrEmpty(rule)) return;

            // Simple parser: assumes format {Param1}-{Param2}
            var parts = rule.Split(new[] { "}{", "{", "}" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var part in parts)
            {
                if (part == TxtSeparator.Text) continue;
                var match = _allParamsSource.FirstOrDefault(p => p.Name == part);
                if (match != null)
                {
                    SelectedParams.Add(match.Name);
                }
            }
            RefreshAvailableList();
        }

        private void RefreshAvailableList()
        {
            AvailableParams.Clear();
            string search = TxtSearch.Text?.ToLower() ?? "";

            int catIndex = ComboSource != null ? ComboSource.SelectedIndex : 0;

            foreach (var p in _allParamsSource)
            {
                // 1. Search Filter
                if (!p.Name.ToLower().Contains(search)) continue;

                // 2. Already Selected Filter
                if (SelectedParams.Contains(p.Name)) continue;

                // 3. Category Filter
                bool categoryMatch = true;
                if (catIndex == 1) // Sheet Parameters
                {
                    categoryMatch = (p.Group == "Sheet Parameter" || p.Group == "Built-in");
                }
                else if (catIndex == 2) // Project Info
                {
                    categoryMatch = (p.Group == "Project Information" || p.Group == "Project Info");
                }

                if (categoryMatch)
                {
                    AvailableParams.Add(p.Name);
                }
            }
        }

        // --- EVENT HANDLERS ---

        private void OnSourceChanged(object sender, SelectionChangedEventArgs e)
        {
            RefreshAvailableList();
        }

        private void OnSearchChanged(object sender, TextChangedEventArgs e)
        {
            RefreshAvailableList();
        }

        private void OnAddClick(object sender, RoutedEventArgs e)
        {
            if (ListAvailable.SelectedItem is string itemToAdd)
            {
                SelectedParams.Add(itemToAdd);
                RefreshAvailableList();
            }
        }

        private void OnRemoveClick(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button btn && btn.DataContext is string itemToRemove)
            {
                SelectedParams.Remove(itemToRemove);
                RefreshAvailableList();
            }
            else if (ListSelected.SelectedItem is string selectedItemToRemove && !string.IsNullOrEmpty(selectedItemToRemove))
            {
                SelectedParams.Remove(selectedItemToRemove);
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

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}