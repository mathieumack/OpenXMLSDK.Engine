using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Platform.Word.Extensions;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine
{
    /// <summary>
    /// Extension class for Table Cell rendering
    /// </summary>
    public static class CellExtensions
    {
        /// <summary>
        /// Render a cell in a word document
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="parent"></param>
        /// <param name="context"></param>
        /// <param name="documentPart"></param>
        /// <returns></returns>
        public static TableCell Render(this Cell cell, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart)
        {
            context.ReplaceItem(cell);

            TableCell wordCell = new TableCell();

            TableCellProperties cellProp = new TableCellProperties();
            wordCell.AppendChild(cellProp);

            // cells properties
            if (cell.Borders != null)
            {
                TableCellBorders borders = cell.Borders.RenderCellBorder();
                cellProp.AppendChild(borders);
            }
            if( !string.IsNullOrEmpty(cell.Shading))
            {
                cellProp.Shading = new Shading() { Fill = cell.Shading };
            }
            if (cell.VerticalAlignment.HasValue)
            {
                cellProp.AppendChild(new TableCellVerticalAlignment { Val = cell.VerticalAlignment.Value.ToOOxml() });
            }
            if(cell.TextDirection.HasValue)
            {
                cellProp.AppendChild(new TextDirection { Val = cell.TextDirection.ToOOxml() });
            }

            // manage cell column and row span
            if(cell.ColSpan > 1)
            {
                cellProp.AppendChild(new GridSpan() { Val = cell.ColSpan });
            }
            if(cell.Fusion)
            {
                if(cell.FusionChild)
                {
                    cellProp.AppendChild(new VerticalMerge() { Val = MergedCellValues.Continue });
                }
                else
                {
                    cellProp.AppendChild(new VerticalMerge() { Val = MergedCellValues.Restart });
                }
            }

            // fill cell content (need at least an empty paragraph)
            var paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
            if (cell.Justification.HasValue)
            {
                var ppr = new ParagraphProperties();
                ppr.AppendChild(new Justification() { Val = cell.Justification.Value.ToOOxml() });
                paragraph.AppendChild(ppr);
            }
            wordCell.AppendChild(paragraph);
            foreach (var element in cell.ChildElements)
            {
                var content = element.Render(paragraph, context, documentPart);
            }

            return wordCell;
        }
    }
}
