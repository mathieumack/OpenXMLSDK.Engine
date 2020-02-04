using DO = DocumentFormat.OpenXml;
using DOP = DocumentFormat.OpenXml.Packaging;
using DOW = DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.Platform.Word.Extensions;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using System;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    public static class PageExtensions
    {
        public static void Render(this Page page, Document document, DO.OpenXmlElement wdDoc, ContextModel context, DOP.MainDocumentPart mainDocumentPart, IFormatProvider formatProvider)
        {
            if (!string.IsNullOrWhiteSpace(page.ShowKey) && context.ExistItem<BooleanModel>(page.ShowKey) && !context.GetItem<BooleanModel>(page.ShowKey).Value)
                return;

            // add page content
            ((BaseElement)page).Render(document, wdDoc, context, mainDocumentPart, formatProvider);

            // add section to manage orientation. Last section is at the end of document
            var pageSize = new DOW.PageSize()
            {
                Orient = page.PageOrientation.ToOOxml(),
                Width = DO.UInt32Value.FromUInt32(page.PageOrientation == PageOrientationValues.Landscape ? (uint)16839 : 11907),
                Height = DO.UInt32Value.FromUInt32(page.PageOrientation == PageOrientationValues.Landscape ? (uint)11907 : 16839)
            };
            var sectionProps = new DOW.SectionProperties(pageSize);
            // document margins
            if (page.Margin != null)
            {
                var pageMargins = new DOW.PageMargin()
                {
                    Left = page.Margin.Left,
                    Top = page.Margin.Top,
                    Right = page.Margin.Right,
                    Bottom = page.Margin.Bottom,
                    Footer = page.Margin.Footer,
                    Header = page.Margin.Header
                };
                sectionProps.AppendChild(pageMargins);
            }
            var p = new DOW.Paragraph();
            var ppr = new DOW.ParagraphProperties();
            p.AppendChild(ppr);
            ppr.AppendChild(sectionProps);
            wdDoc.AppendChild(p);
        }
    }
}
