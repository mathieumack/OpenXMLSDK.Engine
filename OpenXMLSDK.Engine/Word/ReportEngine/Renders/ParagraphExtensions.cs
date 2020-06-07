using System;
using DocumentFormat.OpenXml;
using OpenXMLSDK.Engine.ReportEngine.DataContext;
using OpenXMLSDK.Engine.Word.ReportEngine.Models;
using OpenXMLSDK.Engine.Platform.Word.Extensions;
using OpenXMLSDK.Engine.Word.ReportEngine.Models.ExtendedModels;
using System.Linq;
using System.IO;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    public static class ParagraphExtensions
    {
        public static OpenXmlElement Render(this Paragraph paragraph, OpenXmlElement parent, ContextModel context, IFormatProvider formatProvider)
        {
            context.ReplaceItem(paragraph, formatProvider);

            //if (paragraph.ChildElements.OfType<Paragraph>().Any())
            //    throw new InvalidDataException("A paragraph can not be a direct child of an other paragrah");

            var openXmlPar = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
            openXmlPar.ParagraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties();

            if (!string.IsNullOrWhiteSpace(paragraph.Shading))
                openXmlPar.ParagraphProperties.Shading = new DocumentFormat.OpenXml.Wordprocessing.Shading() { 
                    Fill = paragraph.Shading, 
                    Val = new EnumValue<DocumentFormat.OpenXml.Wordprocessing.ShadingPatternValues>(DocumentFormat.OpenXml.Wordprocessing.ShadingPatternValues.Clear)
                };

            openXmlPar.ParagraphProperties.Justification = new DocumentFormat.OpenXml.Wordprocessing.Justification() { Val = paragraph.Justification.ToOOxml() };
            openXmlPar.ParagraphProperties.SpacingBetweenLines = new DocumentFormat.OpenXml.Wordprocessing.SpacingBetweenLines();

            if (paragraph.SpacingBefore.HasValue)
                openXmlPar.ParagraphProperties.SpacingBetweenLines.Before = paragraph.SpacingBefore.ToString();
            if (paragraph.SpacingAfter.HasValue)
                openXmlPar.ParagraphProperties.SpacingBetweenLines.After = paragraph.SpacingAfter.ToString();
            if (paragraph.SpacingBetweenLines.HasValue)
                openXmlPar.ParagraphProperties.SpacingBetweenLines.Line = paragraph.SpacingBetweenLines.ToString();
            if (!string.IsNullOrWhiteSpace(paragraph.ParagraphStyleId))
                openXmlPar.ParagraphProperties.ParagraphStyleId = new DocumentFormat.OpenXml.Wordprocessing.ParagraphStyleId() { Val = paragraph.ParagraphStyleId };
            if (paragraph.Borders != null)
                openXmlPar.ParagraphProperties.ParagraphBorders = paragraph.Borders.RenderParagraphBorder();
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
            if (!string.IsNullOrWhiteSpace(indentation.Left))
                result.Left = indentation.Left;
            if (indentation.LeftChars.HasValue)
                result.LeftChars = indentation.LeftChars.Value;

            // Right :
            if (!string.IsNullOrWhiteSpace(indentation.Right))
                result.Right = indentation.Right;
            if (indentation.RightChars.HasValue)
                result.RightChars = indentation.RightChars.Value;

            return result;
        }
    }
}
