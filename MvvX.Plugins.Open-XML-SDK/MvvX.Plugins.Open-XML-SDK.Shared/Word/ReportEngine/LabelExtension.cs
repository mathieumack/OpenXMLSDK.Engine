using System;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine
{
    public static class LabelExtension
    {
        /// <summary>
        /// Render a label
        /// </summary>
        /// <param name="label"></param>
        /// <param name="parent"></param>
        /// <param name="context"></param>
        /// <param name="documentPart"></param>
        /// <param name="formatProvider"></param>
        /// <returns></returns>
        public static OpenXmlElement Render(this Label label, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart, IFormatProvider formatProvider)
        {
            context.ReplaceItem(label, formatProvider);

            if (label.IsHtml)
            {
                AlternativeFormatImportPart formatImportPart;
                if (documentPart is MainDocumentPart)
                    formatImportPart = (documentPart as MainDocumentPart).AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Xhtml);
                else if (documentPart is HeaderPart)
                    formatImportPart = (documentPart as HeaderPart).AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Xhtml);
                else if (documentPart is FooterPart)
                    formatImportPart = (documentPart as FooterPart).AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Xhtml);
                else
                    return null;

                return SetHtmlContent(label, parent, documentPart, formatImportPart);
            }
            else
            {
                return SetTextContent(label, parent);
            }
        }

        /// <summary>
        /// Set html content.
        /// </summary>
        /// <param name="label"></param>
        /// <param name="parent"></param>
        /// <param name="documentPart"></param>
        /// <param name="formatImportPart"></param>
        /// <returns></returns>
        private static AltChunk SetHtmlContent(Label label, OpenXmlElement parent, OpenXmlPart documentPart, AlternativeFormatImportPart formatImportPart)
        {
            AltChunk altChunk = new AltChunk();
            altChunk.Id = documentPart.GetIdOfPart(formatImportPart);

            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(label.Text)))
            {
                formatImportPart.FeedData(ms);
            }

            OpenXmlElement paragraph = null;
            if (parent is DocumentFormat.OpenXml.Wordprocessing.Paragraph)
            {
                paragraph = parent;
            }
            else
            {
                paragraph = parent.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault();
            }

            if (paragraph != null)
            {
                paragraph.InsertAfterSelf(altChunk);
            }

            return altChunk;
        }

        /// <summary>
        /// Set text content
        /// </summary>
        /// <param name="label"></param>
        /// <param name="parent"></param>
        /// <returns></returns>
        private static Run SetTextContent(Label label, OpenXmlElement parent)
        {
            Run run = new Run();

            if (label.Text == null)
            {
                run.AppendChild(new Text(label.Text)
                {
                    Space = (SpaceProcessingModeValues)(int)label.SpaceProcessingModeValue
                });
            }
            else
            {
                var lines = label.Text.Split('\n');

                for (int i = 0; i < lines.Length; i++)
                {
                    run.AppendChild(new Text(lines[i])
                    {
                        Space = (SpaceProcessingModeValues)(int)label.SpaceProcessingModeValue
                    });
                    if (i < lines.Length - 1)
                    {
                        run.AppendChild(new Break());
                    }
                }
            }

            var runProperty = new RunProperties();
            if (!string.IsNullOrWhiteSpace(label.FontName))
                runProperty.RunFonts = new RunFonts() { Ascii = label.FontName, HighAnsi = label.FontName, EastAsia = label.FontName, ComplexScript = label.FontName };
            if (!string.IsNullOrWhiteSpace(label.FontSize))
                runProperty.FontSize = new FontSize() { Val = label.FontSize };
            if (!string.IsNullOrWhiteSpace(label.FontSize))
                runProperty.Color = new Color() { Val = label.FontColor };
            if (!string.IsNullOrWhiteSpace(label.Shading))
                runProperty.Shading = new Shading() { Fill = label.Shading };
            if (label.Bold.HasValue)
                runProperty.Bold = new Bold() { Val = OnOffValue.FromBoolean(label.Bold.Value) };
            if (label.Italic.HasValue)
                runProperty.Italic = new Italic() { Val = OnOffValue.FromBoolean(label.Italic.Value) };
            if (label.IsPageNumber)
                run.AppendChild(new PageNumber());

            if (label.Underline != null)
            {
                var underline = new Underline();
                underline.Val = (UnderlineValues)(int)label.Underline.Val;
                if (!string.IsNullOrWhiteSpace(label.Underline.Color))
                    underline.Color = label.Underline.Color;
                runProperty.Underline = underline;
            }

            run.RunProperties = runProperty;
            parent.Append(run);

            return run;
        }
    }
}
