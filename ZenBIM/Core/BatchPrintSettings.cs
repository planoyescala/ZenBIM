//-----------------------------------------------------------------------------------------
// <copyright file="BatchPrintSettings.cs" company="plano y escala">
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
using System.IO;
using System.Text.Json;

namespace ZenBIM.Core
{
    public class BatchPrintSettings
    {
        public string LastNamingRule { get; set; } = "{Sheet Number}-{Sheet Name}";
        public string LastSeparator { get; set; } = "-";

        // Nueva propiedad para guardar la ruta
        public string LastOutputFolder { get; set; } = string.Empty;

        private static string SettingsPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ZenBIM",
            "BatchPrintSettings.json");

        public static void Save(BatchPrintSettings settings)
        {
            try
            {
                string? dir = Path.GetDirectoryName(SettingsPath);
                if (dir != null && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                string json = JsonSerializer.Serialize(settings);
                File.WriteAllText(SettingsPath, json);
            }
            catch { /* Fail silently */ }
        }

        public static BatchPrintSettings Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    string json = File.ReadAllText(SettingsPath);
                    return JsonSerializer.Deserialize<BatchPrintSettings>(json) ?? new BatchPrintSettings();
                }
            }
            catch { /* Fail silently */ }
            return new BatchPrintSettings();
        }
    }
}