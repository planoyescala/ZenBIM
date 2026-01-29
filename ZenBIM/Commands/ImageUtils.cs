//-----------------------------------------------------------------------------------------
// <copyright file="ImageUtils.cs" company="plano y escala">
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
using System.Windows.Media.Imaging;

namespace ZenBIM.Utils
{
    public static class ImageUtils
    {
        public static BitmapSource GetImageSource(string resourcePath)
        {
            // Fix for CS8603: Using null-coalescing to prevent null returns
            var uri = new Uri(resourcePath, UriKind.RelativeOrAbsolute);
            var bitmap = new BitmapImage(uri);
            return bitmap ?? throw new InvalidOperationException("Resource not found");
        }
    }
}