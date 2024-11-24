using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.ReportEngine.DataContext;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    /// <summary>
    /// Extension class for footers
    /// </summary>
    public static class FooterExtensions
    {
        /// <summary>
        /// Render the header of document
        /// </summary>
        /// <param name="footer"></param>
        /// <param name="document"></param>
        /// <param name="mainDocumentPart"></param>
        /// <param name="context"></param>
        /// <param name="formatProvider"></param>
        public static void Render(this Models.Footer footer, Models.Document document, MainDocumentPart mainDocumentPart, ContextModel context, IFormatProvider formatProvider)
        {
            var footerPart = mainDocumentPart.AddNewPart<FooterPart>();

            footerPart.Footer = new Footer();

            foreach (var element in footer.ChildElements)
            {
                element.InheritFromParent(footer);
                element.Render(document, footerPart.Footer, context, footerPart, formatProvider);
            }

            string footerPartId = mainDocumentPart.GetIdOfPart(footerPart);
            if (!mainDocumentPart.Document.Body.Descendants<SectionProperties>().Any())
            {
                mainDocumentPart.Document.Body.AppendChild(new SectionProperties());
            }
            foreach (var section in mainDocumentPart.Document.Body.Descendants<SectionProperties>())
            {
                section.PrependChild(new FooterReference() { Id = footerPartId, Type = new DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues(footer.Type.ToString().ToLower()) });
            }

            if (footer.Type == HeaderFooterValues.First)
            {
                mainDocumentPart.Document.Body.Descendants<SectionProperties>().First().PrependChild(new TitlePage());
            }
        }
    }
}
