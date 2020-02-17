using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.Text;

namespace Pdf.Engine.ReportEngine.Extensions
{
    public static class ParagraphExtensions
    {
        /// <summary>
        /// Apply a style on a paragraph
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="document"></param>
        public static void ApplyStyle(this Paragraph paragraph, Document document)
        {
            // Override values from paragraph style id :
            if (!string.IsNullOrWhiteSpace(paragraph.ParagraphStyleId) && document.Styles.Any(e => e.StyleId.Equals(paragraph.ParagraphStyleId)))
            {
                var style = document.Styles.First(e => e.StyleId.Equals(paragraph.ParagraphStyleId));

                if (!string.IsNullOrWhiteSpace(style.FontName))
                    paragraph.FontName = style.FontName;
                if (!string.IsNullOrWhiteSpace(style.FontColor))
                    paragraph.FontColor = style.FontColor;
                if (!string.IsNullOrWhiteSpace(style.Shading))
                    paragraph.Shading = style.Shading;
                if (style.FontSize.HasValue)
                    paragraph.FontSize = style.FontSize;
                if (!string.IsNullOrWhiteSpace(style.FontEncoding))
                    paragraph.FontColor = style.FontEncoding;
            }
        }
    }
}
