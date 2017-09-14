using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Platform.Word.Extensions;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine
{
    public static class PageExtensions
    {
        public static void Render(this Page page, OpenXmlElement wdDoc, ContextModel context, MainDocumentPart mainDocumentPart)
        {
            if (!string.IsNullOrWhiteSpace(page.ShowKey) && context.ExistItem<BooleanModel>(page.ShowKey) && !context.GetItem<BooleanModel>(page.ShowKey).Value)
                return;

            // add page content
            ((BaseElement)page).Render(wdDoc, context, mainDocumentPart);

            // add section to manage orientation. Last section is at the end of document
            var pageSize = new PageSize()
            {
                Orient = page.PageOrientation.ToOOxml(),
                Width = UInt32Value.FromUInt32(page.PageOrientation == OpenXMLSDK.Word.PageOrientationValues.Landscape ? (uint)16839 : 11907),
                Height = UInt32Value.FromUInt32(page.PageOrientation == OpenXMLSDK.Word.PageOrientationValues.Landscape ? (uint)11907 : 16839)
            };
            var sectionProps = new SectionProperties(pageSize);
            // document margins
            if (page.Margin != null)
            {
                var pageMargins = new PageMargin()
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
            var p = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
            var ppr = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties();
            p.AppendChild(ppr);
            ppr.AppendChild(sectionProps);
            wdDoc.AppendChild(p);
        }
    }
}
