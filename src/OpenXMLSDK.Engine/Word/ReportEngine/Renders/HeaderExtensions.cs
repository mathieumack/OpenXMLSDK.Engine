using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.ReportEngine.DataContext;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
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
        /// <param name="document"></param>
        /// <param name="mainDocumentPart"></param>
        /// <param name="context"></param>
        /// <param name="formatProvider"></param>
        public static void Render(this Models.Header header, Models.Document document, MainDocumentPart mainDocumentPart, ContextModel context, IFormatProvider formatProvider)
        {
            var headerPart = mainDocumentPart.AddNewPart<HeaderPart>();

            headerPart.Header = new Header();

            foreach (var element in header.ChildElements)
            {
                element.InheritFromParent(header);
                element.Render(document, headerPart.Header, context, headerPart, formatProvider);
            }

            string headerPartId = mainDocumentPart.GetIdOfPart(headerPart);
            if (!mainDocumentPart.Document.Body.Descendants<SectionProperties>().Any())
            {
                mainDocumentPart.Document.Body.AppendChild(new SectionProperties());
            }
            foreach (var section in mainDocumentPart.Document.Body.Descendants<SectionProperties>())
            {
                section.PrependChild(new HeaderReference() { Id = headerPartId, Type = (DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues)(int)header.Type });
            }

            if (header.Type == HeaderFooterValues.First)
            {
                mainDocumentPart.Document.Body.Descendants<SectionProperties>().First().PrependChild(new TitlePage());
            }
        }
    }
}
