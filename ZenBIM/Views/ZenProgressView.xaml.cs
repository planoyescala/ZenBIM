//-----------------------------------------------------------------------------------------
// <copyright file="ZenProgressView.xaml.cs" company="plano y escala">
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
using System.Windows.Media;
using System.Windows.Threading;

namespace ZenBIM.Views
{
    public partial class ZenProgressView : Window
    {
        public ZenProgressView(string mainTitle, string subTitle = "Processing...")
        {
            InitializeComponent();
            TxtStatus.Text = mainTitle ?? "ZenBIM: Working";
            TxtSubStatus.Text = subTitle ?? "Please wait";
        }

        public void SetIcon(string pathData)
        {
            if (!string.IsNullOrEmpty(pathData))
                IconPath.Data = Geometry.Parse(pathData);
        }

        public void Update(double value)
        {
            double safeValue = Math.Max(0, Math.Min(100, value));

            // Asignación directa
            ProgBar.Value = safeValue;
            TxtPercent.Text = $"{(int)safeValue}%";

            // Forzar refresco visual
            Dispatcher.Invoke(() => { }, DispatcherPriority.Background);
        }

        protected override void OnMouseDown(System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left) this.DragMove();
        }
    }
}