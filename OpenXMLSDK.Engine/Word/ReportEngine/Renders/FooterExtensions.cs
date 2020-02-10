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
        public static void Render(this Footer footer, Document document, DOP.MainDocumentPart mainDocumentPart, ContextModel context, IFormatProvider formatProvider)
        {
            var footerPart = mainDocumentPart.AddNewPart<DOP.FooterPart>();

            footerPart.Footer = new DOW.Footer();

            foreach (var element in footer.ChildElements)
            {
                element.InheritsFromParent(footer);
                element.Render(document, footerPart.Footer, context, footerPart, formatProvider);
            }

            string footerPartId = mainDocumentPart.GetIdOfPart(footerPart);
            if (!mainDocumentPart.Document.Body.Descendants<DOW.SectionProperties>().Any())
            {
                mainDocumentPart.Document.Body.AppendChild(new DOW.SectionProperties());
            }
            foreach (var section in mainDocumentPart.Document.Body.Descendants<DOW.SectionProperties>())
            {
                section.PrependChild(new DOW.FooterReference() { Id = footerPartId, Type = (DOW.HeaderFooterValues)(int)footer.Type });
            }

            if (footer.Type == HeaderFooterValues.First)
            {
                mainDocumentPart.Document.Body.Descendants<DOW.SectionProperties>().First().PrependChild(new DOW.TitlePage());
            }
        }
    }
}
