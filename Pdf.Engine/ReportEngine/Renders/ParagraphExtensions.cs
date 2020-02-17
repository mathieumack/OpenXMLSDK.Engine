using System;
using it = iTextSharp.text;
using itp = iTextSharp.text.pdf;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.ExtendedModels;
using ReportEngine.Core.Template.Extensions;
using ReportEngine.Core.Template.Text;
using System.Linq;
using Pdf.Engine.ReportEngine.Extensions;

namespace Pdf.Engine.ReportEngine.Renders
{
    public static class ParagraphExtensions
    {
        internal static void Render(this Paragraph paragraph,
                                                    Document document,
                                                    itp.PdfWriter writer,
                                                    it.Document pdfDocument,
                                                    ContextModel context,
                                                    EngineContext ctx,
                                                    IFormatProvider formatProvider)
        {
            context.ReplaceItem(paragraph, formatProvider);

            paragraph.ApplyStyle(document);

            var pdfParagraph = new it.Paragraph();
            //pdfParagraph.Alignment = (int)paragraph.HorizontAlignement;

            //pdfParagraph.SpacingAfter = paragraph.SpacingAfter;
            //pdfParagraph.SpacingBefore = paragraph.SpacingBefore;
            //pdfParagraph.Leading = paragraph.Leading;

            foreach (var child in paragraph.ChildElements)
            {
                child.InheritsFromParent(paragraph);
                var childElement = child.Render(document, writer, pdfDocument, context, ctx, formatProvider);
                pdfParagraph.Add(childElement);
            }

            // Indents :
            if (paragraph.Indentation != null)
            {
                if(paragraph.Indentation.Left.HasValue)
                    pdfParagraph.IndentationLeft = paragraph.Indentation.Left.Value / 20;
                if (paragraph.Indentation.Right.HasValue)
                    pdfParagraph.IndentationRight = paragraph.Indentation.Right.Value / 20;
            }

            // Shading :

            // Borders are not managed yet.

            paragraph.AddToParentContainer(ctx, pdfParagraph);
            //var openXmlPar = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
            //openXmlPar.ParagraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties()
            //{
            //    Shading = new DocumentFormat.OpenXml.Wordprocessing.Shading() { Fill = paragraph.Shading },
            //    Justification = new DocumentFormat.OpenXml.Wordprocessing.Justification() { Val = paragraph.Justification.ToOOxml() },
            //    SpacingBetweenLines = new DocumentFormat.OpenXml.Wordprocessing.SpacingBetweenLines()
            //};
            //if (paragraph.SpacingBefore.HasValue)
            //    openXmlPar.ParagraphProperties.SpacingBetweenLines.Before = paragraph.SpacingBefore.ToString();
            //if (paragraph.SpacingAfter.HasValue)
            //    openXmlPar.ParagraphProperties.SpacingBetweenLines.After = paragraph.SpacingAfter.ToString();
            //if (paragraph.SpacingBetweenLines.HasValue)
            //    openXmlPar.ParagraphProperties.SpacingBetweenLines.Line = paragraph.SpacingBetweenLines.ToString();
            //if (!string.IsNullOrWhiteSpace(paragraph.ParagraphStyleId))
            //    openXmlPar.ParagraphProperties.ParagraphStyleId = new DocumentFormat.OpenXml.Wordprocessing.ParagraphStyleId() { Val = paragraph.ParagraphStyleId };
            //if (paragraph.Borders != null)
            //{
            //    openXmlPar.ParagraphProperties.AppendChild(paragraph.Borders.RenderParagraphBorder());
            //}
            //if (paragraph.Keeplines)
            //    openXmlPar.ParagraphProperties.KeepLines = new DocumentFormat.OpenXml.Wordprocessing.KeepLines();
            //if (paragraph.KeepNext)
            //    openXmlPar.ParagraphProperties.KeepNext = new DocumentFormat.OpenXml.Wordprocessing.KeepNext();
            //if (paragraph.PageBreakBefore)
            //    openXmlPar.ParagraphProperties.PageBreakBefore = new DocumentFormat.OpenXml.Wordprocessing.PageBreakBefore();

            //parent.Append(openXmlPar);
        }

        ///// <summary>
        ///// Transform an indentation model to an OpenXml element
        ///// </summary>
        ///// <param name="indentation"></param>
        ///// <returns></returns>
        //private static DocumentFormat.OpenXml.Wordprocessing.Indentation ToOpenXmlElement(this ParagraphIndentationModel indentation)
        //{
        //    var result = new DocumentFormat.OpenXml.Wordprocessing.Indentation();

        //    // Left :
        //    if (!string.IsNullOrWhiteSpace(indentation.Left))
        //        result.Left = indentation.Left;
        //    if (indentation.LeftChars.HasValue)
        //        result.LeftChars = indentation.LeftChars.Value;

        //    // Right :
        //    if (!string.IsNullOrWhiteSpace(indentation.Right))
        //        result.Right = indentation.Right;
        //    if (indentation.RightChars.HasValue)
        //        result.RightChars = indentation.RightChars.Value;

        //    return result;
        //}
    }
}
