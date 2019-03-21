using DocumentFormat.OpenXml.Packaging;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace OpenXMLSDK.Platform.Word.ReportEngine
{
    public static class DocumentExtensions
    {
        /// <summary>
        /// Render the document
        /// </summary>
        /// <param name="document"></param>
        /// <param name="wdDoc"></param>
        /// <param name="context"></param>
        public static void Render(this Document document, WordprocessingDocument wdDoc, ContextModel context)
        {
            // add styles in document
            var spart = wdDoc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
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
                    ((ForEachPage)pageItem).Render(wdDoc.MainDocumentPart.Document.Body, context, wdDoc.MainDocumentPart, document);
                }
                else if(pageItem is Page)
                {
                    var page = (Page)pageItem;
                    bool addPageBreak = (document.Pages.IndexOf(page) < document.Pages.Count - 1);

                    // doc inherit margin from page
                    if (document.Margin == null && page.Margin != null)
                        document.Margin = page.Margin;
                    // page inherit margin from doc
                    else if (document.Margin != null && page.Margin == null)
                        page.Margin = document.Margin;

                    // render page
                    page.Render(wdDoc.MainDocumentPart.Document.Body, context, wdDoc.MainDocumentPart);
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
                footer.Render(wdDoc.MainDocumentPart, context);
            }
            // headers
            foreach (var header in document.Headers)
            {
                header.Render(wdDoc.MainDocumentPart, context);
            }
        }

        /// <summary>
        /// Render the document
        /// </summary>
        /// <param name="document"></param>
        /// <param name="wdDoc"></param>
        /// <param name="context"></param>
        public static void Render(this Document document, WordprocessingDocument wdDoc, ContextModel context, bool addPageBreak)
        {
            foreach (var pageItem in document.Pages)
            {
                if (pageItem is ForEachPage)
                {
                    // render page
                    ((ForEachPage)pageItem).Render(wdDoc.MainDocumentPart.Document.Body, context, wdDoc.MainDocumentPart, document);
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
                    page.Render(wdDoc.MainDocumentPart.Document.Body, context, wdDoc.MainDocumentPart);
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
