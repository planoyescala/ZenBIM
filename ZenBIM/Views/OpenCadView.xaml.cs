//-----------------------------------------------------------------------------------------
// <copyright file="OpenCadView.xaml.cs" company="plano y escala">
// Copyright (c) 2026 plano y escala.
//
// ZenBIM is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// </copyright>
// <author>plano y escala</author>
//-----------------------------------------------------------------------------------------

using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;

namespace ZenBIM.Views
{
    public partial class OpenCadView : Window
    {
        private string _dwgPath;
        public bool ReloadRequested { get; private set; } = false;

        public OpenCadView(string path)
        {
            InitializeComponent();
            _dwgPath = path;
            TxtPath.Text = path;
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) this.DragMove();
        }

        private void BtnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            if (File.Exists(_dwgPath))
            {
                try
                {
                    // Let Windows handle file association (AutoCAD, TrueView, etc.)
                    Process.Start(new ProcessStartInfo(_dwgPath) { UseShellExecute = true });
                    this.Close();
                }
                catch (System.Exception ex)
                {
                    System.Windows.MessageBox.Show(
                        $"Error opening file:\n{ex.Message}", "ZenBIM Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                System.Windows.MessageBox.Show(
                    "The file does not exist at the saved path.\nIt may have been moved or renamed.",
                    "File Not Found",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
            }
        }

        private void BtnOpenFolder_Click(object sender, RoutedEventArgs e)
        {
            string? folder = Path.GetDirectoryName(_dwgPath);

            if (!string.IsNullOrEmpty(folder) && Directory.Exists(folder))
            {
                Process.Start("explorer.exe", $"\"{folder}\"");
                this.Close();
            }
            else
            {
                System.Windows.MessageBox.Show(
                    "Could not find the containing folder.",
                    "Folder Not Found",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
            }
        }

        private void BtnCopyPath_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Clipboard.SetText(_dwgPath);
            this.Close();
        }

        private void BtnReload_Click(object sender, RoutedEventArgs e)
        {
            ReloadRequested = true;
            this.Close();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}