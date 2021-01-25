using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.Platform.Word.Extensions;
using OpenXMLSDK.Engine.ReportEngine.DataContext;
using OpenXMLSDK.Engine.Word.ReportEngine.Models;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    public static class PageExtensions
    {
        public static void Render(this Page page, Models.Document document, OpenXmlElement wdDoc, ContextModel context, MainDocumentPart mainDocumentPart, IFormatProvider formatProvider)
        {
            if (context.TryGetItem(page.ShowKey, out BooleanModel showPageItem) && !showPageItem.Value)
            {
                return;
            }

            // add page content
            ((BaseElement)page).Render(document, wdDoc, context, mainDocumentPart, formatProvider);

            // add section to manage orientation. Last section is at the end of document
            var pageSize = new PageSize()
            {
                Orient = page.PageOrientation.ToOOxml(),
                Width = UInt32Value.FromUInt32(page.PageOrientation == PageOrientationValues.Landscape ? (uint)16839 : 11907),
                Height = UInt32Value.FromUInt32(page.PageOrientation == PageOrientationValues.Landscape ? (uint)11907 : 16839)
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

            // If a ColumnCount is defined with this split paragraph in columns
            if (page.ColumnCount.HasValue)
            {
                var columns = new DocumentFormat.OpenXml.Wordprocessing.Columns
                {
                    EqualWidth = true,
                    ColumnCount = (Int16)page.ColumnCount.Value
                };

                // Add columns in section
                sectionProps.Append(columns);
            }

            var p = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
            var ppr = new ParagraphProperties();
            p.AppendChild(ppr);
            ppr.AppendChild(sectionProps);
            wdDoc.AppendChild(p);
        }
    }
}
