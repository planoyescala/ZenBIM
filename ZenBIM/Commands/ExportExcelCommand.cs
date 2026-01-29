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
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ZenBIM.Views;

namespace ZenBIM.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ExportExcelCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get UIDocument and Document
            var uidoc = commandData.Application.ActiveUIDocument;
            var doc = uidoc.Document;

            // Initialize and show the window
            ExportExcelView view = new ExportExcelView(doc);
            view.ShowDialog();

            return Result.Succeeded;
        }
    }
}