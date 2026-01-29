//-----------------------------------------------------------------------------------------
// <copyright file="OpenCadCommand.cs" company="plano y escala">
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
using System.Linq;

namespace ZenBIM.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class OpenCadCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // 1. Get Selection
            var selectionIds = uidoc.Selection.GetElementIds();

            if (selectionIds.Count != 1)
            {
                // Explicit namespace used to avoid CS0104 error
                Autodesk.Revit.UI.TaskDialog.Show("ZenBIM", "Please select exactly ONE CAD link (ImportInstance).");
                return Result.Cancelled;
            }

            Element elem = doc.GetElement(selectionIds.First());

            // Validate ImportInstance
            if (!(elem is ImportInstance importInst))
            {
                Autodesk.Revit.UI.TaskDialog.Show("ZenBIM", "The selected element is not a CAD link.");
                return Result.Cancelled;
            }

            // 2. Get File Path
            CADLinkType? cadLinkType = doc.GetElement(importInst.GetTypeId()) as CADLinkType;

            // Check if it is a Link (not an embedded import)
            if (cadLinkType == null || !cadLinkType.IsExternalFileReference())
            {
                Autodesk.Revit.UI.TaskDialog.Show("ZenBIM", "The selected CAD is imported (embedded), not linked. Path cannot be managed.");
                return Result.Failed;
            }

            ExternalFileReference extRef = cadLinkType.GetExternalFileReference();
            string path = ModelPathUtils.ConvertModelPathToUserVisiblePath(extRef.GetAbsolutePath());

            // 3. Show UI
            OpenCadView view = new OpenCadView(path);
            view.ShowDialog();

            // 4. Process Reload if requested
            if (view.ReloadRequested)
            {
                // For CADLinks we DO use a Transaction
                using (Transaction t = new Transaction(doc, "Reload DWG"))
                {
                    t.Start();
                    cadLinkType.Reload();
                    t.Commit();
                }

                Autodesk.Revit.UI.TaskDialog.Show("ZenBIM", "CAD Link reloaded successfully.");
            }

            return Result.Succeeded;
        }
    }
}