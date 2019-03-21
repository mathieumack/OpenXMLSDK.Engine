using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;
using System.Globalization;
using MvvX.Plugins.OpenXMLSDK.Platform.Word.Extensions;

namespace OpenXMLSDK.Platform.Word.ReportEngine
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
            if (!string.IsNullOrEmpty(cell.Shading))
            {
                cellProp.Shading = new Shading() { Fill = cell.Shading };
            }
            if (cell.VerticalAlignment.HasValue)
            {
                cellProp.AppendChild(new TableCellVerticalAlignment { Val = cell.VerticalAlignment.Value.ToOOxml() });
            }
            if (cell.TextDirection.HasValue)
            {
                cellProp.AppendChild(new TextDirection { Val = cell.TextDirection.ToOOxml() });
            }
            if (cell.CellWidth != null)
            {
                cellProp.AppendChild(new TableCellWidth() { Width = cell.CellWidth.Width, Type = cell.CellWidth.Type.ToOOxml() });
            }

            // manage cell column and row span
            if (cell.ColSpan > 1)
            {
                cellProp.AppendChild(new GridSpan() { Val = cell.ColSpan });
            }
            if (cell.Fusion)
            {
                if (cell.FusionChild)
                {
                    cellProp.AppendChild(new VerticalMerge() { Val = MergedCellValues.Continue });
                }
                else
                {
                    cellProp.AppendChild(new VerticalMerge() { Val = MergedCellValues.Restart });
                }
            }
            if (cell.Margin != null)
            {
                cellProp.AppendChild(new TableCellMargin()
                {
                    LeftMargin = new LeftMargin() { Width = cell.Margin.Left.ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa },
                    TopMargin = new TopMargin() { Width = cell.Margin.Top.ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa },
                    RightMargin = new RightMargin() { Width = cell.Margin.Right.ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa },
                    BottomMargin = new BottomMargin() { Width = cell.Margin.Bottom.ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa }
                });
            }

            if (cell.Show)
            {
                if (cell.ChildElements.Any(x => x is MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models.Paragraph)
              || cell.ChildElements.Any(x => x is MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models.ForEach))
                {
                    foreach (var element in cell.ChildElements)
                    {
                        element.InheritFromParent(cell);
                        if (element is MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models.Paragraph
                            || element is MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models.ForEach)
                        {
                            element.Render(wordCell, context, documentPart);
                        }
                        else
                        {
                            var paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
                            if (cell.Justification.HasValue)
                            {
                                var ppr = new ParagraphProperties();
                                ppr.AppendChild(new Justification() { Val = cell.Justification.Value.ToOOxml() });
                                paragraph.AppendChild(ppr);
                            }
                            wordCell.AppendChild(paragraph);
                            var r = new Run();
                            paragraph.AppendChild(r);
                            element.Render(r, context, documentPart);
                        }
                    }
                }
                else
                {
                    // fill cell content (need at least an empty paragraph)
                    var paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
                    if (cell.Justification.HasValue)
                    {
                        var ppr = new ParagraphProperties();
                        ppr.AppendChild(new Justification() { Val = cell.Justification.Value.ToOOxml() });
                        paragraph.AppendChild(ppr);
                    }
                    wordCell.AppendChild(paragraph);
                    var r = new Run();
                    paragraph.AppendChild(r);
                    foreach (var element in cell.ChildElements)
                    {
                        element.InheritFromParent(cell);
                        element.Render(r, context, documentPart);
                    }
                }
            }          

            return wordCell;
        }
    }
}
