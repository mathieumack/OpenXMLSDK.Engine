using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.Word.ReportEngine.BatchModels;

namespace OpenXMLSDK.Engine.Word.ReportEngine
{
    /// <summary>
    /// Extension class for footers
    /// </summary>
    public static class FooterExtensions
    {
        /// <summary>
        /// Render the header of document
        /// </summary>
        /// <param name="header"></param>
        /// <param name="mainDocumentPart"></param>
        /// <param name="context"></param>
        /// <param name="formatProvider"></param>
        public static void Render(this OpenXMLSDK.Word.ReportEngine.Models.Footer footer, MainDocumentPart mainDocumentPart, ContextModel context, IFormatProvider formatProvider)
        {
            var footerPart = mainDocumentPart.AddNewPart<FooterPart>();

            footerPart.Footer = new Footer();

            foreach (var element in footer.ChildElements)
            {
                element.InheritFromParent(footer);
                element.Render(footerPart.Footer, context, footerPart, formatProvider);
            }

            string footerPartId = mainDocumentPart.GetIdOfPart(footerPart);
            if (!mainDocumentPart.Document.Body.Descendants<SectionProperties>().Any())
            {
                mainDocumentPart.Document.Body.AppendChild(new SectionProperties());
            }
            foreach (var section in mainDocumentPart.Document.Body.Descendants<SectionProperties>())
            {
                section.PrependChild(new FooterReference() { Id = footerPartId, Type = (DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues)(int)footer.Type });
            }

            if (footer.Type == HeaderFooterValues.First)
            {
                mainDocumentPart.Document.Body.Descendants<SectionProperties>().First().PrependChild(new TitlePage());
            }
        }
    }
}
