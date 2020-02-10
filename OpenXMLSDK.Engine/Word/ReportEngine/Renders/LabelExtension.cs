using System;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template.Extensions;
using ReportEngine.Core.Template.Text;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    internal static class LabelExtension
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
        internal static OpenXmlElement Render(this Label label, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart, IFormatProvider formatProvider)
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
        /// <param name="formatProvider"></param>
        /// <returns></returns>
        private static Run SetTextContent(Label label, OpenXmlElement parent)
        {
            Run run = new Run();

            // Transform label Text before rendering :
            ApplyTransformOperations(label);

            if (label.Text == null)
            {
                run.AppendChild(new Text(label.Text)
                {
                    Space = (DocumentFormat.OpenXml.SpaceProcessingModeValues)(int)label.SpaceProcessingModeValue
                });
            }
            else
            {
                var lines = label.Text.Split('\n');

                for (int i = 0; i < lines.Length; i++)
                {
                    run.AppendChild(new Text(lines[i])
                    {
                        Space = (DocumentFormat.OpenXml.SpaceProcessingModeValues)(int)label.SpaceProcessingModeValue
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
            if (label.FontSize.HasValue)
                runProperty.FontSize = new FontSize() { Val = label.FontSize.Value.ToString() };
            if (!string.IsNullOrWhiteSpace(label.FontColor))
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
                var underline = new Underline()
                {
                    Val = (UnderlineValues)(int)label.Underline.Val
                };

                if (!string.IsNullOrWhiteSpace(label.Underline.Color))
                    underline.Color = label.Underline.Color;

                runProperty.Underline = underline;
            }

            run.RunProperties = runProperty;
            parent.Append(run);

            return run;
        }

        /// <summary>
        /// Apply transforme operation on the label before rendering
        /// </summary>
        /// <param name="label"></param>
        private static void ApplyTransformOperations(Label label)
        {
            if(!string.IsNullOrWhiteSpace(label.Text) && label.TransformOperations != null)
            {
                foreach(var operation in label.TransformOperations.Where(e => e != null))
                {
                    switch(operation.TransformOperationType)
                    {
                        case LabelTransformOperationType.ToUpper:
                            label.Text = label.Text.ToUpper();
                            break;
                        case LabelTransformOperationType.ToLower:
                            label.Text = label.Text.ToLower();
                            break;
                        case LabelTransformOperationType.ToUpperInvariant:
                            label.Text = label.Text.ToUpperInvariant();
                            break;
                        case LabelTransformOperationType.ToLowerInvariant:
                            label.Text = label.Text.ToLowerInvariant();
                            break;
                        case LabelTransformOperationType.Trim:
                            label.Text = label.Text.Trim();
                            break;
                        case LabelTransformOperationType.TrimStart:
                            label.Text = label.Text.TrimStart();
                            break;
                        case LabelTransformOperationType.TrimEnd:
                            label.Text = label.Text.TrimEnd();
                            break;
                        default:
                            break;
                    }
                }
            }
        }
    }
}
