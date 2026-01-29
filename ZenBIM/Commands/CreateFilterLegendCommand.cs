//-----------------------------------------------------------------------------------------
// <copyright file="CreateFilterLegendCommand.cs" company="plano y escala">
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
using System.Collections.Generic;
using System.Linq;

namespace ZenBIM.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class CreateFilterLegendCommand : IExternalCommand
    {
        private class FilterData
        {
            public string Name { get; }
            public OverrideGraphicSettings Settings { get; }
            public bool IsHalftone { get; }

            public FilterData(string name, OverrideGraphicSettings settings, bool isHalftone)
            {
                Name = name;
                Settings = settings;
                IsHalftone = isHalftone;
            }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            // --- PHASE 1: UI ---
            // Ensure FilterLegendView is also translated in your project if it has UI
            var ui = new FilterLegendView(doc);
            if (ui.ShowDialog() != true) return Result.Cancelled;

            Autodesk.Revit.DB.View? sourceView = ui.SelectedView;
            if (sourceView == null) return Result.Failed;

            // --- PROGRESS BAR START ---
            var progress = new ZenProgressView("ZenBIM: Legends", "Analyzing view filters...");
            // Filter / List Icon
            progress.SetIcon("M10,18H14V16H10V18M3,6V8H21V6H3M6,13H18V11H6V13Z");
            progress.Show();
            // --------------------------

            try
            {
                // --- PHASE 2: DATA GATHERING ---
                List<FilterData> safeFilterList = new List<FilterData>();
                ICollection<ElementId> filterIds = sourceView.GetFilters();

                foreach (ElementId fid in filterIds)
                {
                    Element? elem = doc.GetElement(fid);
                    if (elem == null) continue;

                    OverrideGraphicSettings ogs = sourceView.GetFilterOverrides(fid);
                    bool halftone = ogs.Halftone;

                    safeFilterList.Add(new FilterData(elem.Name, ogs, halftone));
                }

                if (safeFilterList.Count == 0)
                {
                    // Hide progress bar before showing modal message
                    progress.Visibility = System.Windows.Visibility.Hidden;
                    System.Windows.MessageBox.Show("No filters found in this view.", "ZenBIM");
                    return Result.Cancelled;
                }

                // --- PHASE 3: MEASUREMENTS & RESOURCES ---
                FilledRegionType? regionType = new FilteredElementCollector(doc)
                    .OfClass(typeof(FilledRegionType))
                    .Cast<FilledRegionType>()
                    .FirstOrDefault();

                if (regionType == null) { message = "Filled Region Type missing in project."; return Result.Failed; }

                // Convert UI inputs (cm) to Internal Units (feet)
                double wFill = UnitUtils.ConvertToInternalUnits(ui.RegionWidth, UnitTypeId.Centimeters);
                double wLine = UnitUtils.ConvertToInternalUnits(ui.LineWidth, UnitTypeId.Centimeters);
                double h = UnitUtils.ConvertToInternalUnits(ui.RegionHeight, UnitTypeId.Centimeters);

                double gap = UnitUtils.ConvertToInternalUnits(15, UnitTypeId.Centimeters);
                double textPadding = UnitUtils.ConvertToInternalUnits(20, UnitTypeId.Centimeters);
                double spacingY = h * 1.5;

                // X Coordinates calculation
                double xCol1_SurfFill = 0;
                double xCol2_SurfLine = xCol1_SurfFill + wFill + gap;
                double xCol3_CutFill = xCol2_SurfLine + wLine + gap;
                double xCol4_CutLine = xCol3_CutFill + wFill + gap;
                double xText = xCol4_CutLine + wLine + textPadding;

                // --- PHASE 4: TRANSACTION ---
                using (Transaction t = new Transaction(doc, "ZenBIM: Create Legend"))
                {
                    t.Start();
                    try
                    {
                        // Duplicate an existing Legend to use as a canvas
                        var existingLegend = new FilteredElementCollector(doc)
                            .OfClass(typeof(Autodesk.Revit.DB.View))
                            .Cast<Autodesk.Revit.DB.View>()
                            .FirstOrDefault(v => v.ViewType == ViewType.Legend && !v.IsTemplate);

                        if (existingLegend == null)
                        {
                            progress.Visibility = System.Windows.Visibility.Hidden;
                            System.Windows.MessageBox.Show("A base Legend view is required in the project to duplicate.");
                            return Result.Cancelled;
                        }

                        ElementId newId = existingLegend.Duplicate(ViewDuplicateOption.Duplicate);
                        doc.Regenerate();

                        Autodesk.Revit.DB.View legendView = (Autodesk.Revit.DB.View)doc.GetElement(newId);
                        if (legendView.Scale != 100) legendView.Scale = 100;

                        try { legendView.Name = $"Legend_{sourceView.Name}_{DateTime.Now:HHmm}"; } catch { }

                        // DRAW ROWS
                        XYZ origin = XYZ.Zero;
                        int total = safeFilterList.Count;
                        int count = 0;

                        foreach (FilterData data in safeFilterList)
                        {
                            count++;
                            progress.Update(((double)count / total) * 100);

                            OverrideGraphicSettings fs = data.Settings;
                            double y = origin.Y; // Top Y of the row

                            // 1. SURFACE - FILL
                            DrawFill(doc, legendView, regionType.Id, new XYZ(xCol1_SurfFill, y, 0), wFill, h, fs, data.IsHalftone, isCut: false);

                            // 2. SURFACE - LINE
                            DrawLine(doc, legendView, new XYZ(xCol2_SurfLine, y, 0), wLine, h, fs, data.IsHalftone, isCut: false);

                            // 3. CUT - FILL
                            DrawFill(doc, legendView, regionType.Id, new XYZ(xCol3_CutFill, y, 0), wFill, h, fs, data.IsHalftone, isCut: true);

                            // 4. CUT - LINE
                            DrawLine(doc, legendView, new XYZ(xCol4_CutLine, y, 0), wLine, h, fs, data.IsHalftone, isCut: true);

                            // 5. TEXT
                            if (ui.SelectedTextType != null)
                            {
                                // ALIGNMENT FIX
                                double textY = y - (h / 2.0) + (h * 0.08);
                                TextNote.Create(doc, legendView.Id, new XYZ(xText, textY, 0), data.Name, ui.SelectedTextType.Id);
                            }

                            // Move origin down for next row
                            origin = new XYZ(origin.X, origin.Y - spacingY, 0);
                        }

                        t.Commit();
                        commandData.Application.ActiveUIDocument.ActiveView = legendView;
                    }
                    catch (Exception ex)
                    {
                        t.RollBack();
                        message = "Error: " + ex.Message;
                        return Result.Failed;
                    }
                }
            }
            finally
            {
                // Ensure progress bar closes
                progress.Close();
            }

            return Result.Succeeded;
        }

        // --- HELPER METHODS ---

        private void DrawFill(Document doc, Autodesk.Revit.DB.View view, ElementId regionTypeId, XYZ pos, double w, double h, OverrideGraphicSettings settings, bool isHalftone, bool isCut)
        {
            // Determine if Pattern or Color is present
            bool hasPattern = isCut ? settings.CutForegroundPatternId != ElementId.InvalidElementId
                                    : settings.SurfaceForegroundPatternId != ElementId.InvalidElementId;

            bool hasColor = isCut ? settings.CutForegroundPatternColor.IsValid
                                  : settings.SurfaceForegroundPatternColor.IsValid;

            // Check Visibility state
            bool isVisible = isCut ? settings.IsCutForegroundPatternVisible
                                   : settings.IsSurfaceForegroundPatternVisible;

            // CRITICAL FIX: Ghost Fill Prevention
            if (isVisible && !hasPattern && !hasColor) return;

            // Prepare Override
            OverrideGraphicSettings boxOverrides = new OverrideGraphicSettings();

            if (!isVisible)
            {
                // HIDDEN -> Force White Solid to mask
                SetSolidFill(doc, boxOverrides);
                boxOverrides.SetSurfaceForegroundPatternColor(new Autodesk.Revit.DB.Color(255, 255, 255));
            }
            else
            {
                // Standard Configuration (Only if we have overrides)
                ElementId patId = isCut ? settings.CutForegroundPatternId : settings.SurfaceForegroundPatternId;
                Autodesk.Revit.DB.Color col = isCut ? settings.CutForegroundPatternColor : settings.SurfaceForegroundPatternColor;

                if (patId != ElementId.InvalidElementId)
                {
                    boxOverrides.SetSurfaceForegroundPatternId(patId);
                    boxOverrides.SetSurfaceForegroundPatternColor(col);
                }
                else if (col.IsValid)
                {
                    SetSolidFill(doc, boxOverrides);
                    boxOverrides.SetSurfaceForegroundPatternColor(col);
                }
            }

            // Create Geometry (Rectangle)
            List<CurveLoop> loops = new List<CurveLoop>();
            CurveLoop loop = new CurveLoop();
            XYZ p1 = pos;
            XYZ p2 = new XYZ(pos.X + w, pos.Y, 0);
            XYZ p3 = new XYZ(pos.X + w, pos.Y - h, 0);
            XYZ p4 = new XYZ(pos.X, pos.Y - h, 0);

            loop.Append(Line.CreateBound(p1, p2));
            loop.Append(Line.CreateBound(p2, p3));
            loop.Append(Line.CreateBound(p3, p4));
            loop.Append(Line.CreateBound(p4, p1));
            loops.Add(loop);

            try
            {
                FilledRegion region = FilledRegion.Create(doc, regionTypeId, view.Id, loops);

                if (isHalftone) boxOverrides.SetHalftone(true);

                // Make borders invisible (White)
                boxOverrides.SetProjectionLineColor(new Autodesk.Revit.DB.Color(255, 255, 255));

                view.SetElementOverrides(region.Id, boxOverrides);
            }
            catch
            {
                // Ignore geometry errors for very small loops
            }
        }

        private void DrawLine(Document doc, Autodesk.Revit.DB.View view, XYZ pos, double length, double cellHeight, OverrideGraphicSettings settings, bool isHalftone, bool isCut)
        {
            Autodesk.Revit.DB.Color col = isCut ? settings.CutLineColor : settings.ProjectionLineColor;
            ElementId patId = isCut ? settings.CutLinePatternId : settings.ProjectionLinePatternId;
            int weight = isCut ? settings.CutLineWeight : settings.ProjectionLineWeight;

            // If no properties, skip
            if (!col.IsValid && patId == ElementId.InvalidElementId && weight == -1) return;

            // Coordinates: Center vertically in the cell
            double centerY = pos.Y - (cellHeight / 2.0);
            XYZ start = new XYZ(pos.X, centerY, 0);
            XYZ end = new XYZ(pos.X + length, centerY, 0);

            DetailCurve dc = doc.Create.NewDetailCurve(view, Line.CreateBound(start, end));

            OverrideGraphicSettings lineOverrides = new OverrideGraphicSettings();
            if (col.IsValid) lineOverrides.SetProjectionLineColor(col);
            if (patId != ElementId.InvalidElementId) lineOverrides.SetProjectionLinePatternId(patId);
            if (weight != -1) lineOverrides.SetProjectionLineWeight(weight);

            if (isHalftone) lineOverrides.SetHalftone(true);

            view.SetElementOverrides(dc.Id, lineOverrides);
        }

        private void SetSolidFill(Document doc, OverrideGraphicSettings ogs)
        {
            FillPatternElement? solid = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(p => p.GetFillPattern().IsSolidFill);
            if (solid != null) ogs.SetSurfaceForegroundPatternId(solid.Id);
        }
    }
}