//-----------------------------------------------------------------------------------------
// <copyright file="ExportExcelCommand.cs" company="plano y escala">
// Copyright (c) 2026 plano y escala.
//
// ZenBIM is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// </copyright>
// <author>plano y escala</author>
//-----------------------------------------------------------------------------------------

using Autodesk.Revit.Attributes;
using Nice3point.Revit.Toolkit.External;
using System.Diagnostics;

namespace ZenBIM.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class DonateCommand : ExternalCommand
    {
        public override void Execute()
        {
            // Opens the default web browser with the specified URL
            Process.Start(new ProcessStartInfo
            {
                FileName = "https://buymeacoffee.com/planoyescala",
                UseShellExecute = true
            });
        }
    }
}