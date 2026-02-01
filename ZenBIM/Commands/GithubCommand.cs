//-----------------------------------------------------------------------------------------
// <copyright file="GithubCommand.cs" company="plano y escala">
// Copyright (c) 2026 plano y escala.
//
// ZenBIM is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// </copyright>
// <author>plano y escala</author>
//-----------------------------------------------------------------------------------------

using System.Diagnostics;
using Autodesk.Revit.Attributes;
using Nice3point.Revit.Toolkit.External;

namespace ZenBIM.Commands
{
    [UsedImplicitly]
    [Transaction(TransactionMode.Manual)]
    public class GithubCommand : ExternalCommand
    {
        public override void Execute()
        {
            // Abre la web del proyecto
            Process.Start(new ProcessStartInfo("https://github.com/planoyescala/ZenBIM") { UseShellExecute = true });
        }
    }
}