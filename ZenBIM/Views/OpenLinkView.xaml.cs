//-----------------------------------------------------------------------------------------
// <copyright file="OpenLinkView.xaml.cs" company="plano y escala">
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
    public partial class OpenLinkView : Window
    {
        private string _rvtPath;
        public bool ReloadRequested { get; private set; } = false;

        public OpenLinkView(string path)
        {
            InitializeComponent();
            _rvtPath = path;
            TxtPath.Text = path;
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) this.DragMove();
        }

        private void BtnOpenModel_Click(object sender, RoutedEventArgs e)
        {
            if (File.Exists(_rvtPath))
            {
                try
                {
                    // ROBUST LAUNCH STRATEGY:
                    // 1. Find the current Revit executable.
                    var currentProcess = Process.GetCurrentProcess();
                    if (currentProcess.MainModule == null)
                    {
                        throw new System.Exception("Could not locate Revit executable.");
                    }

                    string revitExePath = currentProcess.MainModule.FileName;
                    string arguments = $"\"{_rvtPath}\"";

                    // 2. Explicitly launch new instance
                    Process.Start(revitExePath, arguments);

                    this.Close();
                }
                catch (System.Exception ex)
                {
                    System.Windows.MessageBox.Show(
                        $"Error launching Revit:\n{ex.Message}",
                        "Launch Error",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            }
            else
            {
                // File not found (Cloud Model or broken path)
                System.Windows.MessageBox.Show(
                    $"File not found at:\n{_rvtPath}\n\nNote: Cloud models (BIM 360/ACC) cannot be opened directly from local paths.",
                    "File Inaccessible",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }

        private void BtnOpenFolder_Click(object sender, RoutedEventArgs e)
        {
            string? folder = Path.GetDirectoryName(_rvtPath);

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
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }

        private void BtnCopyPath_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Clipboard.SetText(_rvtPath);
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