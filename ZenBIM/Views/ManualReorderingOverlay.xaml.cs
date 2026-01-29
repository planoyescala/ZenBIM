//-----------------------------------------------------------------------------------------
// <copyright file="ManualReorderingOverlay.xaml.cs" company="plano y escala">
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
using System.Windows;
using System.Windows.Input;

namespace ZenBIM.Views
{
    public partial class ManualReorderingOverlay : Window
    {
        public event EventHandler? UndoRequested;
        public event EventHandler? ResumeRequested;
        public event EventHandler? FinishRequested;

        public ManualReorderingOverlay()
        {
            InitializeComponent();
        }

        public void UpdateValue(string val)
        {
            TxtNextValue.Text = val;
        }

        public void SetStatus(string text, bool isPaused)
        {
            TxtStatus.Text = text;

            if (isPaused)
            {
                // PAUSED Visual Config
                TxtStatus.Foreground = System.Windows.Media.Brushes.Orange;

                BtnUndo.IsEnabled = true;

                BtnAction.Content = "▶ Continue";
                BtnAction.IsEnabled = true;

                // Explicit ZenBIM Green
                var color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#00C853");
                BtnAction.Background = new System.Windows.Media.SolidColorBrush(color);
            }
            else
            {
                // SELECTING Visual Config
                TxtStatus.Foreground = System.Windows.Media.Brushes.Gray;

                BtnUndo.IsEnabled = false;

                BtnAction.Content = "Selecting...";
                BtnAction.IsEnabled = false;
                BtnAction.Background = System.Windows.Media.Brushes.LightGray;
            }
        }

        private void BtnUndo_Click(object sender, RoutedEventArgs e)
        {
            UndoRequested?.Invoke(this, EventArgs.Empty);
        }

        private void BtnAction_Click(object sender, RoutedEventArgs e)
        {
            ResumeRequested?.Invoke(this, EventArgs.Empty);
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            FinishRequested?.Invoke(this, EventArgs.Empty);
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) DragMove();
        }
    }
}