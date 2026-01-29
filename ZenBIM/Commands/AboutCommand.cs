//-----------------------------------------------------------------------------------------
// <copyright file="AboutCommand.cs" company="plano y escala">
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

namespace ZenBIM.Commands
{
    [UsedImplicitly]
    [Transaction(TransactionMode.Manual)]
    public class AboutCommand : ExternalCommand
    {
        public override void Execute()
        {
            // Welcome message requested for the GitHub project
            string message = "Welcome to ZenBIM by plano y escala.\n\n" +
                             "This project is Open Source software, licensed under the GNU General Public License v3 (GPL v3). " +
                             "You are free to use, modify, and redistribute it following these terms.\n\n" +
                             "Any ideas, suggestions, or issues are more than welcome at: manuel.planoyescala@gmail.com\n\n" +
                             "Thank you for supporting free software!";

            Autodesk.Revit.UI.TaskDialog.Show("About ZenBIM", message);
        }
    }
}