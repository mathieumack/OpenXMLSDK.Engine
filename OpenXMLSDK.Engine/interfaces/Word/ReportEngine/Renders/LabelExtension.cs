using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.interfaces.Word.ReportEngine.Models;
using OpenXMLSDK.Engine.Word.ReportEngine.BatchModels;
using OpenXMLSDK.Engine.Word.ReportEngine.Models;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
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

            return SetTextContent(label, parent, formatProvider);
        }

        /// <summary>
        /// Set text content
        /// </summary>
        /// <param name="label"></param>
        /// <param name="parent"></param>
        /// <param name="formatProvider"></param>
        /// <returns></returns>
        private static Run SetTextContent(Label label, OpenXmlElement parent, IFormatProvider formatProvider)
        {
            Run run = new Run();

            // Transform label Text before rendering :
            ApplyTransformOperations(label, formatProvider);

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
            if (!string.IsNullOrWhiteSpace(label.FontSize))
                runProperty.FontSize = new FontSize() { Val = label.FontSize };
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

        /// <summary>
        /// Apply transforme operation on the label before rendering
        /// </summary>
        /// <param name="label"></param>
        private static void ApplyTransformOperations(Label label, IFormatProvider formatProvider)
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
