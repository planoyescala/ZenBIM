//-----------------------------------------------------------------------------------------
// <copyright file="ReorderingCommand.cs" company="plano y escala">
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

namespace ZenBIM.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ReorderingCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Initialize the window
            var view = new ReorderingView(commandData);

            // CRITICAL FIX: Use ShowDialog() instead of Show().
            // This opens the window as "Modal", blocking Revit until closed.
            // This is required for the "Auto Apply" button (Transactions) to work without
            // needing a complex ExternalEvent handler for the auto-mode.
            view.ShowDialog();

            return Result.Succeeded;
        }
    }
}