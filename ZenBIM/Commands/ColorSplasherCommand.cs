//-----------------------------------------------------------------------------------------
// <copyright file="ColorSplasherCommand.cs" company="plano y escala">
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
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ZenBIM.Views;
using System;

namespace ZenBIM.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ColorSplasherCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Open the user interface
                // We pass commandData so the View can create its own transactions
                var view = new ColorSplasherView(commandData);
                view.ShowDialog();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                // If something fails before opening the window
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}