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
                if (document.Pages.Count > 1 && document.Pages.IndexOf(page) > 0)
                {
                    // add page break
                    wdDoc.MainDocumentPart.Document.Body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Break() { Type = DocumentFormat.OpenXml.Wordprocessing.BreakValues.Page })));
                }
                // render page
                page.Render(wdDoc.MainDocumentPart.Document.Body, context);
            }
        }
    }
}
