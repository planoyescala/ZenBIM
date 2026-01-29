//-----------------------------------------------------------------------------------------
// <copyright file="ManualRenumberHandler.cs" company="plano y escala">
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
using System.Collections.Generic;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using ZenBIM.Views;

// CORRECCIÓN: Namespace cambiado a ZenBIM.Commands para solucionar el error CS0246
namespace ZenBIM.Commands
{
    public class ManualRenumberHandler : IExternalEventHandler
    {
        // Nullability fixed
        public Window MainWindow { get; set; } = default!;
        public Category TargetCategory { get; set; } = default!;
        public Parameter TargetParameter { get; set; } = default!;
        public string Prefix { get; set; } = string.Empty;
        public string Suffix { get; set; } = string.Empty;
        public int StartNumber { get; set; }
        public int Step { get; set; }
        public bool PadZeros { get; set; }
        public int PadLength { get; set; }

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;

            if (MainWindow != null) MainWindow.Hide();

            ManualReorderingOverlay overlay = new ManualReorderingOverlay();
            // Position overlay in bottom right corner
            overlay.Left = SystemParameters.WorkArea.Width - 400;
            overlay.Top = SystemParameters.WorkArea.Height - 350;
            overlay.Show();

            int currentNum = StartNumber;
            bool keepRunning = true;
            bool isPaused = false;

            Stack<ElementId> historyIds = new Stack<ElementId>();
            Stack<string> historyOldValues = new Stack<string>();
            Stack<int> historyNums = new Stack<int>();

            // --- UI EVENTS ---
            overlay.FinishRequested += (s, e) => {
                keepRunning = false;
                isPaused = false;
            };

            overlay.ResumeRequested += (s, e) => {
                isPaused = false;
            };

            overlay.UndoRequested += (s, e) =>
            {
                if (historyIds.Count > 0)
                {
                    ElementId lastId = historyIds.Pop();
                    string oldVal = historyOldValues.Pop();
                    int lastNum = historyNums.Pop();

                    using (Transaction t = new Transaction(doc, "ZenBIM: Undo"))
                    {
                        t.Start();
                        Element? el = doc.GetElement(lastId);
                        el?.LookupParameter(TargetParameter.Definition.Name)?.Set(oldVal);
                        t.Commit();
                    }
                    currentNum = lastNum;
                    overlay.UpdateValue(GetNextVal(currentNum));

                    overlay.SetStatus("Undone. Press Continue.", true);
                }
            };

            overlay.UpdateValue(GetNextVal(currentNum));

            // --- MAIN LOOP ---
            while (keepRunning)
            {
                try
                {
                    overlay.SetStatus("Select element... (ESC for menu)", false);
                    System.Windows.Forms.Application.DoEvents();

                    Reference r = uidoc.Selection.PickObject(ObjectType.Element, new CategorySelectionFilter(TargetCategory.Id), "Select to number (ESC to Pause)");
                    Element elem = doc.GetElement(r);

                    using (Transaction t = new Transaction(doc, "ZenBIM: Manual"))
                    {
                        t.Start();
                        Parameter? p = elem.LookupParameter(TargetParameter.Definition.Name);
                        if (p != null)
                        {
                            historyIds.Push(elem.Id);
                            historyOldValues.Push(GetParamValue(elem, p));
                            historyNums.Push(currentNum);

                            string newVal = GetNextVal(currentNum);
                            if (p.StorageType == StorageType.String) p.Set(newVal);
                            else if (p.StorageType == StorageType.Integer && int.TryParse(newVal, out int iVal)) p.Set(iVal);
                        }
                        t.Commit();
                    }

                    currentNum += Step;
                    overlay.UpdateValue(GetNextVal(currentNum));
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    isPaused = true;
                    overlay.SetStatus("⏸ PAUSED", true);

                    while (isPaused && keepRunning)
                    {
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("Error: " + ex.Message);
                    keepRunning = false;
                }
            }

            overlay.Close();
            if (MainWindow != null) MainWindow.Show();
        }

        public string GetName() => "ZenBIM Manual Reorder";

        private string GetNextVal(int num)
        {
            string s = num.ToString();
            if (PadZeros) s = s.PadLeft(PadLength, '0');
            return $"{Prefix}{s}{Suffix}";
        }

        private string GetParamValue(Element elem, Parameter p)
        {
            if (p == null) return "";
            if (p.StorageType == StorageType.String) return p.AsString() ?? "";
            return p.AsValueString() ?? "";
        }

        public class CategorySelectionFilter : ISelectionFilter
        {
            private ElementId _catId;
            public CategorySelectionFilter(ElementId catId) { _catId = catId; }
            public bool AllowElement(Element elem)
            {
                if (elem?.Category == null) return false;
                return elem.Category.Id.Value == _catId.Value;
            }
            public bool AllowReference(Reference reference, XYZ position) => false;
        }
    }
}