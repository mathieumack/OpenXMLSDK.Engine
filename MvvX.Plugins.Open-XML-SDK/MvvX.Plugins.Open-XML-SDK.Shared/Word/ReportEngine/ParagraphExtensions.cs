using DocumentFormat.OpenXml;
using MvvX.Plugins.OpenXMLSDK.Platform.Word.Extensions;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;
using System;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine
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
            parent.Append(openXmlPar);
            return openXmlPar;
        }
    }
}
