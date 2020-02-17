using it = iTextSharp.text;
using itp = iTextSharp.text.pdf;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using System;
using Pdf.Engine.ReportEngine.Helpers;
using ReportEngine.Core.Template.Extensions;

namespace Pdf.Engine.ReportEngine.Renders
{
    public static class DocumentExtensions
    {
        /// <summary>
        /// Render the document
        /// </summary>
        /// <param name="document"></param>
        /// <param name="wdDoc"></param>
        /// <param name="context"></param>
        /// <param name="formatProvider"></param>
        internal static void Render(this Document document, itp.PdfWriter pdfWriter, it.Document pdfDocument, ContextModel context, EngineContext ctx, IFormatProvider formatProvider)
        {
            ctx.Parents.Add(document);

            var pageEvent = new PageEventHelper(document.Headers, document.Footers, ctx, context);
            pdfWriter.PageEvent = pageEvent;
            pdfDocument.Open();

            // Association de l'élément EODoc à la racine du document IText
            ctx.IElementContainers.Add(document, pdfDocument); 

            bool FistPage = true;
            foreach (var pageItem in document.Pages)
            {
                if (pageItem is ForEachPage)
                {
                    // render page
                    //((ForEachPage)pageItem).Render(document, wdDoc.MainDocumentPart.Document.Body, context, wdDoc.MainDocumentPart, formatProvider);
                }
                else if (pageItem is Page)
                {
                    var page = (Page)pageItem;
                    page.InheritsFromParent(document);

                    pdfDocument.NewPage();

                    //// doc inherit margin from page
                    //if (document.Margin == null && page.Margin != null)
                    //    document.Margin = page.Margin;
                    //// page inherit margin from doc
                    //else if (document.Margin != null && page.Margin == null)
                    //    page.Margin = document.Margin;

                    // render page
                    page.Render(document, pdfWriter, pdfDocument, context, ctx, formatProvider);
                }
            }
            ctx.Parents.RemoveAt(ctx.Parents.Count - 1);

            //Replace Last page break
            //if (wdDoc.MainDocumentPart.Document.Body.LastChild != null && 
            //    wdDoc.MainDocumentPart.Document.Body.LastChild is DocumentFormat.OpenXml.Wordprocessing.Paragraph &&
            //    wdDoc.MainDocumentPart.Document.Body.LastChild.FirstChild != null &&
            //    wdDoc.MainDocumentPart.Document.Body.LastChild.FirstChild is DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties &&
            //    wdDoc.MainDocumentPart.Document.Body.LastChild.FirstChild.FirstChild != null &&
            //    wdDoc.MainDocumentPart.Document.Body.LastChild.FirstChild.FirstChild is DocumentFormat.OpenXml.Wordprocessing.SectionProperties)
            //{
            //    DocumentFormat.OpenXml.Wordprocessing.Paragraph lastChild = (DocumentFormat.OpenXml.Wordprocessing.Paragraph)wdDoc.MainDocumentPart.Document.Body.LastChild;
            //    DocumentFormat.OpenXml.Wordprocessing.SectionProperties sectionPropertie = (DocumentFormat.OpenXml.Wordprocessing.SectionProperties)lastChild.FirstChild.FirstChild.Clone();
            //    wdDoc.MainDocumentPart.Document.Body.ReplaceChild(sectionPropertie, wdDoc.MainDocumentPart.Document.Body.LastChild);
            //}

            //// footers
            //foreach (var footer in document.Footers)
            //{
            //    footer.Render(document, wdDoc.MainDocumentPart, context, formatProvider);
            //}
            //// headers
            //foreach (var header in document.Headers)
            //{
            //    header.Render(document, wdDoc.MainDocumentPart, context, formatProvider);
            //}
        }
    }
}
