using System.Linq;
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

            foreach (var page in document.Pages)
            {
                bool addPageBreak = (document.Pages.IndexOf(page) < document.Pages.Count -1);

                // page inherit margin from doc
                if(document.Margin != null && page.Margin == null)
                {
                    page.Margin = document.Margin;
                }

                // render page
                page.Render(wdDoc.MainDocumentPart.Document.Body, context, wdDoc.MainDocumentPart, addPageBreak);
            }

            // footer
            if (document.Footer != null)
            {
                document.Footer.Render(wdDoc.MainDocumentPart, context);
            }
            // header
            if (document.Header != null)
            {
                document.Header.Render(wdDoc.MainDocumentPart, context);
            }
        }
    }
}
