using System;
using System.Linq;
using DOP = DocumentFormat.OpenXml.Packaging;
using DOW = DocumentFormat.OpenXml.Wordprocessing;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.Extensions;

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
        public static void Render(this Header header, Document document, DOP.MainDocumentPart mainDocumentPart, ContextModel context, IFormatProvider formatProvider)
        {
            var headerPart = mainDocumentPart.AddNewPart<DOP.HeaderPart>();

            headerPart.Header = new DOW.Header();

            foreach (var element in header.ChildElements)
            {
                element.InheritsFromParent(header);
                element.Render(document, headerPart.Header, context, headerPart, formatProvider);
            }

            string headerPartId = mainDocumentPart.GetIdOfPart(headerPart);
            if (!mainDocumentPart.Document.Body.Descendants<DOW.SectionProperties>().Any())
            {
                mainDocumentPart.Document.Body.AppendChild(new DOW.SectionProperties());
            }
            foreach (var section in mainDocumentPart.Document.Body.Descendants<DOW.SectionProperties>())
            {
                section.PrependChild(new DOW.HeaderReference() { Id = headerPartId, Type = (DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues)(int)header.Type });
            }

            if (header.Type == HeaderFooterValues.First)
            {
                mainDocumentPart.Document.Body.Descendants<DOW.SectionProperties>().First().PrependChild(new DOW.TitlePage());
            }
        }
    }
}
