using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Word.ReportEngine.BatchModels;

namespace OpenXMLSDK.Platform.Word.ReportEngine
{
    /// <summary>
    /// Extension class for rendering header
    /// </summary>
    public static class HeaderExtensions
    {
        /// <summary>
        /// Render the header of document
        /// </summary>
        /// <param name="header"></param>
        /// <param name="mainDocumentPart"></param>
        /// <param name="context"></param>
        public static void Render(this OpenXMLSDK.Word.ReportEngine.Models.Header header, MainDocumentPart mainDocumentPart, ContextModel context)
        {
            var headerPart = mainDocumentPart.AddNewPart<HeaderPart>();

            headerPart.Header = new Header();

            foreach (var element in header.ChildElements)
            {
                element.InheritFromParent(header);
                element.Render(headerPart.Header, context, headerPart);
            }

            string headerPartId = mainDocumentPart.GetIdOfPart(headerPart);
            if (!mainDocumentPart.Document.Body.Descendants<SectionProperties>().Any())
            {
                mainDocumentPart.Document.Body.AppendChild(new SectionProperties());
            }
            foreach (var section in mainDocumentPart.Document.Body.Descendants<SectionProperties>())
            {
                section.PrependChild(new HeaderReference() { Id = headerPartId, Type = (HeaderFooterValues)(int)header.Type });
            }

            if (header.Type == OpenXMLSDK.Word.HeaderFooterValues.First)
            {
                mainDocumentPart.Document.Body.Descendants<SectionProperties>().First().PrependChild(new TitlePage());
            }
        }
    }
}
