using DOP = DocumentFormat.OpenXml.Packaging;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using System;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
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
        public static void Render(this Document document, DOP.WordprocessingDocument wdDoc, ContextModel context, IFormatProvider formatProvider)
        {
            // add styles in document
            var spart = wdDoc.MainDocumentPart.AddNewPart<DOP.StyleDefinitionsPart>();
            spart.Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            foreach (var style in document.Styles)
            {
                style.Render(spart, context);
            }

            foreach (var pageItem in document.Pages)
            {
                if (pageItem is ForEachPage)
                {
                    // render page
                    ((ForEachPage)pageItem).Render(document, wdDoc.MainDocumentPart.Document.Body, context, wdDoc.MainDocumentPart, formatProvider);
                }
                else if(pageItem is Page)
                {
                    var page = (Page)pageItem;

                    // doc inherit margin from page
                    if (document.Margin == null && page.Margin != null)
                        document.Margin = page.Margin;
                    // page inherit margin from doc
                    else if (document.Margin != null && page.Margin == null)
                        page.Margin = document.Margin;

                    // render page
                    page.Render(document, wdDoc.MainDocumentPart.Document.Body, context, wdDoc.MainDocumentPart, formatProvider);
                }
            }

            //Replace Last page break
            if (wdDoc.MainDocumentPart.Document.Body.LastChild != null && 
                wdDoc.MainDocumentPart.Document.Body.LastChild is DocumentFormat.OpenXml.Wordprocessing.Paragraph &&
                wdDoc.MainDocumentPart.Document.Body.LastChild.FirstChild != null &&
                wdDoc.MainDocumentPart.Document.Body.LastChild.FirstChild is DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties &&
                wdDoc.MainDocumentPart.Document.Body.LastChild.FirstChild.FirstChild != null &&
                wdDoc.MainDocumentPart.Document.Body.LastChild.FirstChild.FirstChild is DocumentFormat.OpenXml.Wordprocessing.SectionProperties)
            {
                DocumentFormat.OpenXml.Wordprocessing.Paragraph lastChild = (DocumentFormat.OpenXml.Wordprocessing.Paragraph)wdDoc.MainDocumentPart.Document.Body.LastChild;
                DocumentFormat.OpenXml.Wordprocessing.SectionProperties sectionPropertie = (DocumentFormat.OpenXml.Wordprocessing.SectionProperties)lastChild.FirstChild.FirstChild.Clone();
                wdDoc.MainDocumentPart.Document.Body.ReplaceChild(sectionPropertie, wdDoc.MainDocumentPart.Document.Body.LastChild);
            }

            // footers
            foreach (var footer in document.Footers)
            {
                footer.Render(document, wdDoc.MainDocumentPart, context, formatProvider);
            }
            // headers
            foreach (var header in document.Headers)
            {
                header.Render(document, wdDoc.MainDocumentPart, context, formatProvider);
            }
        }

        /// <summary>
        /// Render the document
        /// </summary>
        /// <param name="document"></param>
        /// <param name="wdDoc"></param>
        /// <param name="context"></param>
        /// <param name="formatProvider"></param>
        public static void Render(this Document document, DOP.WordprocessingDocument wdDoc, ContextModel context, bool addPageBreak, IFormatProvider formatProvider)
        {
            foreach (var pageItem in document.Pages)
            {
                if (pageItem is ForEachPage)
                {
                    // render page
                    ((ForEachPage)pageItem).Render(document, wdDoc.MainDocumentPart.Document.Body, context, wdDoc.MainDocumentPart, formatProvider);
                }
                else if (pageItem is Page)
                {
                    var page = (Page)pageItem;
                   
                    // doc inherit margin from page
                    if (document.Margin == null && page.Margin != null)
                        document.Margin = page.Margin;
                    // page inherit margin from doc
                    else if (document.Margin != null && page.Margin == null)
                        page.Margin = document.Margin;

                    // render page
                    page.Render(document, wdDoc.MainDocumentPart.Document.Body, context, wdDoc.MainDocumentPart, formatProvider);
                }
            }
            //Replace Last page break
            if (!addPageBreak &&
                wdDoc.MainDocumentPart.Document.Body.LastChild != null &&
                wdDoc.MainDocumentPart.Document.Body.LastChild is DocumentFormat.OpenXml.Wordprocessing.Paragraph &&
                wdDoc.MainDocumentPart.Document.Body.LastChild.FirstChild != null &&
                wdDoc.MainDocumentPart.Document.Body.LastChild.FirstChild is DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties &&
                wdDoc.MainDocumentPart.Document.Body.LastChild.FirstChild.FirstChild != null &&
                wdDoc.MainDocumentPart.Document.Body.LastChild.FirstChild.FirstChild is DocumentFormat.OpenXml.Wordprocessing.SectionProperties)
            {
                DocumentFormat.OpenXml.Wordprocessing.Paragraph lastChild = (DocumentFormat.OpenXml.Wordprocessing.Paragraph)wdDoc.MainDocumentPart.Document.Body.LastChild;
                wdDoc.MainDocumentPart.Document.Body.RemoveChild(lastChild);
            }
        }
    }
}
