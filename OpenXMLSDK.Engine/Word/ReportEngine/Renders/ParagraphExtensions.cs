using System;
using DocumentFormat.OpenXml;
using OpenXMLSDK.Engine.Platform.Word.Extensions;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template.ExtendedModels;
using ReportEngine.Core.Template.Extensions;
using ReportEngine.Core.Template.Text;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    public static class ParagraphExtensions
    {
        public static OpenXmlElement Render(this Paragraph paragraph, OpenXmlElement parent, ContextModel context, IFormatProvider formatProvider)
        {
            context.ReplaceItem(paragraph, formatProvider);

            var openXmlPar = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
            openXmlPar.ParagraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties()
            {
                Shading = new DocumentFormat.OpenXml.Wordprocessing.Shading() { Fill = paragraph.Shading },
                Justification = new DocumentFormat.OpenXml.Wordprocessing.Justification() { Val = paragraph.Justification.ToOOxml() },
                SpacingBetweenLines = new DocumentFormat.OpenXml.Wordprocessing.SpacingBetweenLines()
            };
            if (paragraph.SpacingBefore.HasValue)
                openXmlPar.ParagraphProperties.SpacingBetweenLines.Before = paragraph.SpacingBefore.ToString();
            if (paragraph.SpacingAfter.HasValue)
                openXmlPar.ParagraphProperties.SpacingBetweenLines.After = paragraph.SpacingAfter.ToString();
            if (paragraph.SpacingBetweenLines.HasValue)
                openXmlPar.ParagraphProperties.SpacingBetweenLines.Line = paragraph.SpacingBetweenLines.ToString();
            if (!string.IsNullOrWhiteSpace(paragraph.ParagraphStyleId))
                openXmlPar.ParagraphProperties.ParagraphStyleId = new DocumentFormat.OpenXml.Wordprocessing.ParagraphStyleId() { Val = paragraph.ParagraphStyleId };
            if (paragraph.Borders != null)
            {
                openXmlPar.ParagraphProperties.AppendChild(paragraph.Borders.RenderParagraphBorder());
            }
            if (paragraph.Keeplines)
                openXmlPar.ParagraphProperties.KeepLines = new DocumentFormat.OpenXml.Wordprocessing.KeepLines();
            if (paragraph.KeepNext)
                openXmlPar.ParagraphProperties.KeepNext = new DocumentFormat.OpenXml.Wordprocessing.KeepNext();
            if (paragraph.PageBreakBefore)
                openXmlPar.ParagraphProperties.PageBreakBefore = new DocumentFormat.OpenXml.Wordprocessing.PageBreakBefore();

            // Indents :
            if (paragraph.Indentation != null)
                openXmlPar.ParagraphProperties.Indentation = paragraph.Indentation.ToOpenXmlElement();

            parent.Append(openXmlPar);
            return openXmlPar;
        }

        /// <summary>
        /// Transform an indentation model to an OpenXml element
        /// </summary>
        /// <param name="indentation"></param>
        /// <returns></returns>
        private static DocumentFormat.OpenXml.Wordprocessing.Indentation ToOpenXmlElement(this ParagraphIndentationModel indentation)
        {
            var result = new DocumentFormat.OpenXml.Wordprocessing.Indentation();

            // Left :
            if (indentation.Left.HasValue)
                result.Left = indentation.Left.Value.ToString();
            if (indentation.LeftChars.HasValue)
                result.LeftChars = indentation.LeftChars.Value;

            // Right :
            if (indentation.Right.HasValue)
                result.Right = indentation.Right.Value.ToString();
            if (indentation.RightChars.HasValue)
                result.RightChars = indentation.RightChars.Value;

            return result;
        }
    }
}
