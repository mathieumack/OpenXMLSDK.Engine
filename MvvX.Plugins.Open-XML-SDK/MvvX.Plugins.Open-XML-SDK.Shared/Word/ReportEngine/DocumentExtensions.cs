using DocumentFormat.OpenXml.Packaging;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine
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
                    page.Render(wdDoc.MainDocumentPart.Document.Body, context, wdDoc.MainDocumentPart, addPageBreak);
                }
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
    }
}
