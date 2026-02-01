//-----------------------------------------------------------------------------------------
// <copyright file="CheckUpdateCommand.cs" company="plano y escala">
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
using Nice3point.Revit.Toolkit.External;
using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ZenBIM.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class CheckUpdateCommand : ExternalCommand
    {
        private const string RepoOwner = "planoyescala";
        private const string RepoName = "ZenBIM";

        // --- MANUAL MODE (Clicked via Button) ---
        public override void Execute()
        {
            // Manual check: Always shows result (Update found OR Up-to-date)
            RunAutoCheckAsync(silentMode: false).GetAwaiter().GetResult();
        }

        // --- AUTOMATIC MODE (Async Task) ---
        public static async Task RunAutoCheckAsync(bool silentMode = true)
        {
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                // 1. Get Data from GitHub
                GitHubReleaseInfo? releaseInfo = await GetLatestReleaseInfoAsync();

                // 2. Get Local Version
                Version currentVersion = Assembly.GetExecutingAssembly().GetName().Version ?? new Version(1, 0, 0);

                // --- BACK ON UI THREAD ---

                if (releaseInfo == null)
                {
                    // If manual, show error. If silent, ignore connection errors.
                    if (!silentMode) Autodesk.Revit.UI.TaskDialog.Show("ZenBIM Update", "Could not retrieve release info.");
                    return;
                }

                // CHECK IF UPDATE IS AVAILABLE
                if (releaseInfo.Version > currentVersion)
                {
                    if (!silentMode)
                    {
                        // MANUAL MODE: Always show the dialog
                        ShowUpdateDialog(currentVersion, releaseInfo.Version, releaseInfo.TagName);
                    }
                    else
                    {
                        // AUTOMATIC MODE: 
                        // Only show if 7 days have passed since last notification.
                        if (ShouldNotifyUser())
                        {
                            ShowUpdateDialog(currentVersion, releaseInfo.Version, releaseInfo.TagName);

                            // Save today's date so we don't ask again for a week
                            SaveNotificationDate();
                        }
                    }
                }
                else
                {
                    // UP TO DATE
                    if (!silentMode)
                    {
                        // MANUAL: User clicked the button, so we confirm they are up to date.
                        Autodesk.Revit.UI.TaskDialog.Show("ZenBIM Update", $"You are up to date!\n\nInstalled: v{currentVersion}\nLatest: v{releaseInfo.Version}");
                    }
                    else
                    {
                        // AUTOMATIC (STARTUP): 
                        // Do nothing. The user is up to date, so we don't disturb them.
                    }
                }
            }
            catch (Exception ex)
            {
                if (!silentMode)
                {
                    Autodesk.Revit.UI.TaskDialog.Show("ZenBIM Error", "Error checking for updates:\n" + ex.Message);
                }
            }
        }

        // --- PERSISTENCE LOGIC (7-DAY RULE) ---

        private static string GetConfigPath()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folder = Path.Combine(appData, "ZenBIM");

            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }

            return Path.Combine(folder, "last_update_check.txt");
        }

        private static bool ShouldNotifyUser()
        {
            try
            {
                string path = GetConfigPath();
                if (!File.Exists(path)) return true; // Never notified before, so show it.

                string content = File.ReadAllText(path);
                if (long.TryParse(content, out long ticks))
                {
                    DateTime lastCheck = new DateTime(ticks);
                    TimeSpan difference = DateTime.Now - lastCheck;

                    // Return TRUE only if more than 7 days passed
                    return difference.TotalDays >= 7;
                }
            }
            catch
            {
                return true;
            }
            return true;
        }

        private static void SaveNotificationDate()
        {
            try
            {
                string path = GetConfigPath();
                File.WriteAllText(path, DateTime.Now.Ticks.ToString());
            }
            catch { /* Ignore */ }
        }

        // --- GITHUB HELPERS ---

        private class GitHubReleaseInfo
        {
            public Version Version { get; set; } = new Version(0, 0, 0);
            public string TagName { get; set; } = string.Empty;
        }

        private static async Task<GitHubReleaseInfo?> GetLatestReleaseInfoAsync()
        {
            var url = $"https://api.github.com/repos/{RepoOwner}/{RepoName}/releases/latest";

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("User-Agent", "ZenBIM-Updater");
                client.DefaultRequestHeaders.Add("Accept", "application/vnd.github.v3+json");
                client.Timeout = TimeSpan.FromSeconds(5);

                try
                {
                    var response = await client.GetStringAsync(url);

                    if (string.IsNullOrEmpty(response)) return null;

                    var match = Regex.Match(response, "\"tag_name\"\\s*:\\s*\"(.*?)\"");
                    if (!match.Success) return null;

                    string tagName = match.Groups[1].Value;
                    string cleanVersion = tagName.TrimStart('v', 'V');

                    if (Version.TryParse(cleanVersion, out Version? parsedVersion) && parsedVersion != null)
                    {
                        return new GitHubReleaseInfo { Version = parsedVersion, TagName = tagName };
                    }
                    return null;
                }
                catch { return null; }
            }
        }

        private static void ShowUpdateDialog(Version current, Version latest, string tagName)
        {
            Autodesk.Revit.UI.TaskDialog dialog = new Autodesk.Revit.UI.TaskDialog("New Update Available");
            dialog.MainInstruction = "A new version of ZenBIM is available!";
            dialog.MainContent = $"Current Version: {current}\nNew Version: {latest}\n\nDo you want to download it now?";
            dialog.CommonButtons = Autodesk.Revit.UI.TaskDialogCommonButtons.Yes | Autodesk.Revit.UI.TaskDialogCommonButtons.No;
            dialog.DefaultButton = Autodesk.Revit.UI.TaskDialogResult.Yes;

            var result = dialog.Show();

            if (result == Autodesk.Revit.UI.TaskDialogResult.Yes)
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = $"https://github.com/{RepoOwner}/{RepoName}/releases/tag/{tagName}",
                    UseShellExecute = true
                });
            }
        }
    }
}