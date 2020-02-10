using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Pdf.Engine.ReportEngine.Helpers;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.Extensions;
using ReportEngine.Core.Template.Tables;
using ReportEngine.Core.Template.Text;
using it = iTextSharp.text;
using itp = iTextSharp.text.pdf;

namespace Pdf.Engine.ReportEngine.Renders
{
    public static class LabelExtensions
    {
        internal static it.Chunk Render(this Label label,
                                                Document document,
                                                itp.PdfWriter writer,
                                                ContextModel context,
                                                EngineContext ctx,
                                                IFormatProvider formatProvider)
        {
            context.ReplaceItem(label, formatProvider);

            // Transform label Text before rendering :
            ApplyTransformOperations(label);

            // TODO : Manage split lines ?
            var text = label.Text;

            var chunk = new it.Chunk(text, label.GetFont(document.DefaultFontSize));

            //if (label.Underline != null)
            //    chunk.SetUnderline(label.FontUnderline.Thickness, label.FontUnderline.Distance);

            var shading = label.Shading;

            if(string.IsNullOrWhiteSpace(shading))
            {
                // try to check inherits :
                var parent = ctx.Parents.LastOrDefault(e => e is Paragraph || e is Cell);
                if(parent != null)
            }

            if (label.Shading != null)
                chunk.SetBackground(FontHelper.ConverPdfColorToColor(label.Shading));
            
            return chunk;
        }

        public static itp.BaseFont GetBaseFont(this Label element)
        {
            if (element.FontName == FontNames.TIMES_C_ROMAN)
            {
                byte[] file = Ressources.fonts.SimSun;
                return itp.BaseFont.CreateFont("SimSum.ttf", itp.BaseFont.IDENTITY_H, itp.BaseFont.NOT_EMBEDDED, itp.BaseFont.CACHED, file, null);
            }
            else if (element.FontName == FontNames.ARIAL)
            {
                byte[] file = Ressources.fonts.arial;
                return itp.BaseFont.CreateFont("arial.ttf", itp.BaseFont.IDENTITY_H, itp.BaseFont.NOT_EMBEDDED, itp.BaseFont.CACHED, file, null);
            }
            else
                return itp.BaseFont.CreateFont(element.FontName, element.FontEncoding, false);

        }

        public static it.Font GetFont(this Label element, int defaultFontSize)
        {
            itp.BaseFont bf;
            if (element.FontName == FontNames.TIMES_C_ROMAN)
            {
                byte[] file = Ressources.fonts.SimSun;
                bf = itp.BaseFont.CreateFont("SimSum.ttf", itp.BaseFont.IDENTITY_H, itp.BaseFont.NOT_EMBEDDED, itp.BaseFont.CACHED, file, null);
            }
            else if (element.FontName == FontNames.ARIAL)
            {
                byte[] file = Ressources.fonts.arial;
                bf = itp.BaseFont.CreateFont("arial.ttf", itp.BaseFont.IDENTITY_H, itp.BaseFont.NOT_EMBEDDED, itp.BaseFont.CACHED, file, null);
            }
            else
            {
                bf = itp.BaseFont.CreateFont(string.IsNullOrWhiteSpace(element.FontName) ? FontNames.TIMES_ROMAN : element.FontName, 
                                                string.IsNullOrWhiteSpace(element.FontEncoding) ? FontEncodings.CP1252 : element.FontEncoding, 
                                                false);
            }

            // TODO : Add FontSize value check
            var fontStyle = GetFontStyle(element); // Normal
            var font = new it.Font(bf, (element.FontSize.HasValue ? element.FontSize.Value : defaultFontSize) / 2, fontStyle, FontHelper.ConverPdfColorToColor(element.FontColor));
            return font;
        }

        private static int GetFontStyle(Label label)
        {
            //NORMAL = 0,
            //BOLD = 1,
            //ITALIC = 2,
            //BOLDITALIC = 3,
            //UNDERLINE = 4,
            //BOLDUNDERLINE = 5,
            //ITALICUNDERLINE = 6,
            //BOLDITALICUNDERLINE = 7,
            //STRIKETHRU = 8
            return 0; // normal
        }

        /// <summary>
        /// Apply transforme operation on the label before rendering
        /// </summary>
        /// <param name="label"></param>
        private static void ApplyTransformOperations(Label label)
        {
            if (!string.IsNullOrWhiteSpace(label.Text) && label.TransformOperations != null)
            {
                foreach (var operation in label.TransformOperations.Where(e => e != null))
                {
                    switch (operation.TransformOperationType)
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
