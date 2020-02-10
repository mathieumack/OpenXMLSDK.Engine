using System;
using System.Collections.Generic;
using System.Text;

namespace ReportEngine.Core.Template.Extensions
{
    public static class BaseElementExtensions
    {
        public static void InheritsFromParent(this BaseElement element, BaseElement parent)
        {
            if (element is null || parent is null)
                return;

            if (string.IsNullOrEmpty(element.FontColor))
                element.FontColor = parent.FontColor;
            if (string.IsNullOrEmpty(element.FontName))
                element.FontName = parent.FontName;
            if (!element.FontSize.HasValue)
                element.FontSize = parent.FontSize;
            if (string.IsNullOrEmpty(element.Shading))
                element.Shading = parent.Shading;
            if (!element.Bold.HasValue)
                element.Bold = parent.Bold;
            if (!element.Italic.HasValue)
                element.Italic = parent.Italic;
        }
    }
}
