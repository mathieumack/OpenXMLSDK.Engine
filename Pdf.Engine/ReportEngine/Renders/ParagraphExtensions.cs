using System;
using it = iTextSharp.text;
using itp = iTextSharp.text.pdf;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.Extensions;
using ReportEngine.Core.Template.Text;
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
            pdfParagraph.Alignment = paragraph.Justification.ToPdfJustification();

            if(paragraph.SpacingAfter.HasValue)
                pdfParagraph.SpacingAfter = paragraph.SpacingAfter.Value / 20f;
            else
                pdfParagraph.SpacingAfter = 8f; // Default value
            if (paragraph.SpacingBefore.HasValue)
                pdfParagraph.SpacingBefore = paragraph.SpacingBefore.Value / 20f;
            //pdfParagraph.Leading = paragraph.Leading;

            foreach (var child in paragraph.ChildElements)
            {
                child.InheritsFromParent(paragraph);
                var childElement = child.Render(document, writer, pdfDocument, context, ctx, formatProvider);
                if(childElement != null)
                    pdfParagraph.Add(childElement);
            }

            // Indents :
            if (paragraph.Indentation != null)
            {
                if(paragraph.Indentation.Left.HasValue)
                    pdfParagraph.IndentationLeft = paragraph.Indentation.Left.Value / 20f;
                if (paragraph.Indentation.Right.HasValue)
                    pdfParagraph.IndentationRight = paragraph.Indentation.Right.Value / 20f;
            }

            // Shading :

            // Borders are not managed yet.

            paragraph.AddToParentContainer(ctx, pdfParagraph);
        }
    }
}
