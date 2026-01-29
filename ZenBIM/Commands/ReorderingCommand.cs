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
using Autodesk.Revit.UI;
using Nice3point.Revit.Toolkit.External;
using ZenBIM.Views;

namespace ZenBIM.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ReorderingCommand : ExternalCommand
    {
        public override void Execute()
        {
            var view = new ReorderingView(ExternalCommandData);
            view.Show();
        }
    }
}