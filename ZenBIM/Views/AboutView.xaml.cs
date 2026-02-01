//-----------------------------------------------------------------------------------------
// <copyright file="AboutView.xaml.cs" company="plano y escala">
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
using System.Windows;
using System.Windows.Input;
using System.Windows.Navigation;

namespace ZenBIM.Views
{
    public partial class AboutView : Window
    {
        public AboutView()
        {
            InitializeComponent();
        }

        // Allows dragging the window by clicking anywhere on the background
        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        // Closes the window
        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        // Handles the click on the link to open the default browser
        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
            e.Handled = true;
        }
    }
}