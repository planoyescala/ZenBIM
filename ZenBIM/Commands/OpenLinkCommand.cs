//-----------------------------------------------------------------------------------------
// <copyright file="OpenLinkCommand.cs" company="plano y escala">
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
using System.Linq;

namespace ZenBIM.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class OpenLinkCommand : IExternalCommand
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
                Autodesk.Revit.UI.TaskDialog.Show("ZenBIM", "Please select exactly ONE Revit link.");
                return Result.Cancelled;
            }

            Element elem = doc.GetElement(selectionIds.First());

            // Validate if it is a RevitLinkInstance
            if (!(elem is RevitLinkInstance linkInstance))
            {
                Autodesk.Revit.UI.TaskDialog.Show("ZenBIM", "The selected element is not a Revit Link.");
                return Result.Cancelled;
            }

            // 2. Get File Path from Type
            RevitLinkType? linkType = doc.GetElement(linkInstance.GetTypeId()) as RevitLinkType;

            if (linkType == null || !linkType.IsExternalFileReference())
            {
                Autodesk.Revit.UI.TaskDialog.Show("ZenBIM", "Could not retrieve link type information.");
                return Result.Failed;
            }

            // Get Reference
            ExternalFileReference extRef = linkType.GetExternalFileReference();
            string userPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(extRef.GetAbsolutePath());

            // 3. Show UI
            OpenLinkView view = new OpenLinkView(userPath);
            bool? dialogResult = view.ShowDialog();

            // 4. Process Reload if requested
            if (view.ReloadRequested)
            {
                // CRITICAL NOTE: DO NOT OPEN A MANUAL TRANSACTION HERE.
                // The Unload/Reload methods for Revit Links manage their own transactions/regeneration.

                try
                {
                    // 1. Unload (Releases the file)
                    linkType.Unload(null);

                    // 2. Reload (Forces read from disk)
                    LinkLoadResult result = linkType.Reload();

                    if (result.LoadResult == LinkLoadResultType.LinkLoaded)
                    {
                        Autodesk.Revit.UI.TaskDialog.Show("ZenBIM", "Link reloaded successfully! ✅");
                    }
                    else
                    {
                        Autodesk.Revit.UI.TaskDialog.Show("Error", $"Failed to reload link.\nStatus: {result.LoadResult}");
                        return Result.Failed;
                    }
                }
                catch (Exception ex)
                {
                    Autodesk.Revit.UI.TaskDialog.Show("Error", $"Exception during reload:\n{ex.Message}");
                    return Result.Failed;
                }
            }

            return Result.Succeeded;
        }
    }
}