//-----------------------------------------------------------------------------------------
// <copyright file="ScheduleItem.cs" company="plano y escala">
// Copyright (c) 2026 plano y escala.
//
// ZenBIM is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// </copyright>
// <author>plano y escala</author>
//-----------------------------------------------------------------------------------------

using Autodesk.Revit.DB;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ZenBIM.Core
{
    // Data class representing a Revit Schedule with a selection checkbox
    public class ScheduleItem : INotifyPropertyChanged
    {
        // "null!" prevents compiler warning CS8618. We guarantee initialization.
        public ViewSchedule RevitView { get; set; } = null!;
        public string ViewName { get; set; } = string.Empty;
        public string FilterLevel { get; set; } = string.Empty;

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                OnPropertyChanged();
            }
        }

        public ScheduleItem(ViewSchedule schedule, string levelParam)
        {
            RevitView = schedule;
            ViewName = schedule.Name ?? "Unnamed View";

            // Handle potential null parameter
            FilterLevel = string.IsNullOrEmpty(levelParam) ? "Unclassified" : levelParam;
            IsSelected = false;
        }

        // Fix for CS8612/CS8618 (Nullable reference types in Events)
        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}