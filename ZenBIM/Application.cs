//-----------------------------------------------------------------------------------------
// <copyright file="Application.cs" company="plano y escala">
// Copyright (c) 2026 plano y escala.
//
// ZenBIM is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// </copyright>
// <author>plano y escala</author>
//-----------------------------------------------------------------------------------------

using Nice3point.Revit.Toolkit.External;
using ZenBIM.Commands;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.Attributes;
using System;
using System.Reflection; // Required for Assembly location
using System.Threading.Tasks;
using System.Windows.Media.Imaging; // Required for BitmapImage in StackedItems

namespace ZenBIM
{
    [UsedImplicitly]
    public class Application : ExternalApplication
    {
        public override void OnStartup()
        {
            // 1. Create Panels
            var panelPlano = Application.CreatePanel("plano y escala", "ZenBIM");
            var panelDrawings = Application.CreatePanel("Drawing Set", "ZenBIM");
            var panelLinks = Application.CreatePanel("Link Manager", "ZenBIM");
            var panelSchedule = Application.CreatePanel("Schedule", "ZenBIM");
            var panelMiscellaneous = Application.CreatePanel("Miscellaneous", "ZenBIM");

            // ====================================================
            // PANEL 1: plano y escala (Layout: Large + Stacked)
            // ====================================================

            // 1. Large Button: About
            panelPlano.AddPushButton<AboutCommand>("About")
                .SetLargeImage(GetResourcePath("about_32.png"))
                .SetImage(GetResourcePath("about_16.png"))
                .SetToolTip("Information about ZenBIM.");

            // 2. Stacked Column: Update, Support & GitHub
            string assemblyPath = Assembly.GetExecutingAssembly().Location;

            // Create Data for "Update"
            var updateData = new PushButtonData("cmdUpdate", "Update", assemblyPath, typeof(CheckUpdateCommand).FullName)
            {
                ToolTip = "Check for ZenBIM updates.",
                Image = new BitmapImage(new Uri(GetResourcePath("update_16.png"), UriKind.Relative))
            };

            // Create Data for "Support"
            var supportData = new PushButtonData("cmdSupport", "Support", assemblyPath, typeof(DonateCommand).FullName)
            {
                ToolTip = "Support plano y escala on Buy Me a Coffee.",
                Image = new BitmapImage(new Uri(GetResourcePath("Coffe_16.png"), UriKind.Relative))
            };

            // Create Data for "GitHub"
            var githubData = new PushButtonData("cmdGithub", "Repo", assemblyPath, typeof(GithubCommand).FullName)
            {
                ToolTip = "Visit ZenBIM on GitHub.",
                Image = new BitmapImage(new Uri(GetResourcePath("github_16.png"), UriKind.Relative))
            };

            // Add them as a stack of 3 (Max allowed by Revit API)
            panelPlano.AddStackedItems(updateData, supportData, githubData);

            // ====================================================
            // PANEL 2: Drawing Set
            // ====================================================
            panelDrawings.AddPushButton<BatchPrintCommand>("Batch Print")
                .SetLargeImage(GetResourcePath("BatchPrint_32.png"))
                .SetImage(GetResourcePath("BatchPrint_16.png"))
                .SetToolTip("Export selected sheets.");

            panelDrawings.AddSeparator();

            panelDrawings.AddPushButton<CreateFilterLegendCommand>("Filter Legend")
                .SetLargeImage(GetResourcePath("legend_32.png"))
                .SetImage(GetResourcePath("legend_16.png"))
                .SetToolTip("Generate a legend based on view filters.");

            panelDrawings.AddPushButton<ColorSplasherCommand>("Color Splasher")
                .SetLargeImage(GetResourcePath("color_splasher_32.png"))
                .SetImage(GetResourcePath("color_splasher_16.png"))
                .SetToolTip("Colorize elements by parameter values and create legends.");

            // ====================================================
            // PANEL 3: Link Manager
            // ====================================================
            panelLinks.AddPushButton<OpenCadCommand>("Manage DWG")
                .SetLargeImage(GetResourcePath("dwg_32.png"))
                .SetImage(GetResourcePath("dwg_16.png"))
                .SetToolTip("Manage selected CAD link: Open File, Open Folder, or Reload.");

            panelLinks.AddPushButton<OpenLinkCommand>("Manage RVT")
                .SetLargeImage(GetResourcePath("rvt_link_32.png"))
                .SetImage(GetResourcePath("rvt_link_16.png"))
                .SetToolTip("Manage selected Revit Link: Open in new instance, Folder or Reload.");

            // ====================================================
            // PANEL 4: Schedule (Excel Tools)
            // ====================================================
            panelSchedule.AddPushButton<ImportExcelCommand>("Import Excel")
                .SetLargeImage(GetResourcePath("schedule_import_32.png"))
                .SetImage(GetResourcePath("schedule_import_16.png"))
                .SetToolTip("Import data from Excel into a Revit Schedule Header.");

            panelSchedule.AddPushButton<ExportExcelCommand>("Export to Excel")
                .SetLargeImage(GetResourcePath("schedule_export_32.png"))
                .SetImage(GetResourcePath("schedule_export_16.png"))
                .SetToolTip("Export selected Schedules to Excel using Fenix Template.");

            // ====================================================
            // PANEL 5: Miscellaneous
            // ====================================================
            panelMiscellaneous.AddPushButton<ReorderingCommand>("Reorder")
                .SetLargeImage(GetResourcePath("reorder_32.png"))
                .SetImage(GetResourcePath("reorder_16.png"))
                .SetToolTip("Advanced element renumbering tool.");

            // ====================================================
            // IDLING EVENT
            // ====================================================
            this.Application.Idling += OnIdling;
        }

        // Async void to allow awaiting the task on the UI thread
        private async void OnIdling(object? sender, IdlingEventArgs e)
        {
            this.Application.Idling -= OnIdling; // Unsubscribe immediately

            // PRODUCTION MODE: silentMode: true
            // This ensures NO window appears if the user is up-to-date.
            await CheckUpdateCommand.RunAutoCheckAsync(silentMode: true);
        }

        private string GetResourcePath(string iconName)
        {
            return $"/ZenBIM;component/Resources/Icons/{iconName}";
        }

        public override void OnShutdown()
        {
            this.Application.Idling -= OnIdling;
        }
    }
}