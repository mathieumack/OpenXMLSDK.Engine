using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.Word.ReportEngine.BatchModels;
using OpenXMLSDK.Engine.Word.ReportEngine.Models;
using System.Globalization;
using OpenXMLSDK.Engine.Platform.Word.Extensions;
using System;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
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
        /// <param name="document"></param>
        /// <param name="parent"></param>
        /// <param name="context"></param>
        /// <param name="documentPart"></param>
        /// <param name="isInAlternateRow"></param>
        /// <param name="formatProvider"></param>
        /// <returns></returns>
        public static TableCell Render(this Cell cell, 
                                        Models.Document document, 
                                        OpenXmlElement parent, 
                                        ContextModel context, 
                                        OpenXmlPart documentPart, 
                                        bool isInAlternateRow,
                                        IFormatProvider formatProvider)
        {
            context.ReplaceItem(cell, formatProvider);

            TableCell wordCell = new TableCell();

            TableCellProperties cellProp = new TableCellProperties();
            wordCell.AppendChild(cellProp);

            cellProp.AddBorders(cell, isInAlternateRow);
            cellProp.AddShading(cell, isInAlternateRow);
            cellProp.AddMargin(cell, isInAlternateRow);
            cellProp.AddJustifications(cell, isInAlternateRow);

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

            if (cell.Show)
            {
                if (cell.ChildElements.Any(x => x is Models.Paragraph)
              || cell.ChildElements.Any(x => x is ForEach))
                {
                    foreach (var element in cell.ChildElements)
                    {
                        element.InheritFromParent(cell);
                        if (element is Models.Paragraph
                            || element is ForEach)
                        {
                            element.Render(document, wordCell, context, documentPart, formatProvider);
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
                            element.Render(document, r, context, documentPart, formatProvider);
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
                        element.Render(document, r, context, documentPart, formatProvider);
                    }
                }
            }          

            return wordCell;
        }

        /// <summary>
        /// Add border values to the cell properties
        /// </summary>
        /// <param name="cellProp"></param>
        /// <param name="cell"></param>
        /// <param name="isInAlternateRow"></param>
        private static void AddJustifications(this TableCellProperties cellProp, Cell cell, bool isInAlternateRow)
        {
            if (isInAlternateRow && cell.AlternateRowCellConfiguration != null && cell.AlternateRowCellConfiguration.VerticalAlignment.HasValue)
                cellProp.AppendChild(new TableCellVerticalAlignment { Val = cell.AlternateRowCellConfiguration.VerticalAlignment.Value.ToOOxml() });
            else if (cell.VerticalAlignment.HasValue)
                cellProp.AppendChild(new TableCellVerticalAlignment { Val = cell.VerticalAlignment.Value.ToOOxml() });

            if (isInAlternateRow && cell.AlternateRowCellConfiguration != null && cell.AlternateRowCellConfiguration.TextDirection.HasValue)
                cellProp.AppendChild(new TextDirection { Val = cell.AlternateRowCellConfiguration.TextDirection.ToOOxml() });
            else if (cell.TextDirection.HasValue)
                cellProp.AppendChild(new TextDirection { Val = cell.TextDirection.ToOOxml() });

            if (isInAlternateRow && cell.AlternateRowCellConfiguration != null && cell.AlternateRowCellConfiguration.CellWidth != null)
                cellProp.AppendChild(new TableCellWidth() { Width = cell.AlternateRowCellConfiguration.CellWidth.Width, Type = cell.AlternateRowCellConfiguration.CellWidth.Type.ToOOxml() });
            else if (cell.CellWidth != null)
                cellProp.AppendChild(new TableCellWidth() { Width = cell.CellWidth.Width, Type = cell.CellWidth.Type.ToOOxml() });
        }

        /// <summary>
        /// Add shading value to the cell properties
        /// </summary>
        /// <param name="cellProp"></param>
        /// <param name="cell"></param>
        /// <param name="isInAlternateRow"></param>
        private static void AddShading(this TableCellProperties cellProp, Cell cell, bool isInAlternateRow)
        {
            if(isInAlternateRow && cell.AlternateRowCellConfiguration != null && !string.IsNullOrWhiteSpace(cell.AlternateRowCellConfiguration.Shading))
                cellProp.Shading = new Shading() { Fill = cell.AlternateRowCellConfiguration.Shading };
            else if (!string.IsNullOrEmpty(cell.Shading))
                cellProp.Shading = new Shading() { Fill = cell.Shading };
        }

        /// <summary>
        /// Add border values to the cell properties
        /// </summary>
        /// <param name="cellProp"></param>
        /// <param name="cell"></param>
        /// <param name="isInAlternateRow"></param>
        private static void AddBorders(this TableCellProperties cellProp, Cell cell, bool isInAlternateRow)
        {
            if (isInAlternateRow && cell.AlternateRowCellConfiguration != null && cell.AlternateRowCellConfiguration.Borders != null)
            {
                TableCellBorders borders = cell.AlternateRowCellConfiguration.Borders.RenderCellBorder();
                cellProp.AppendChild(borders);
            }
            else if (cell.Borders != null)
            {
                TableCellBorders borders = cell.Borders.RenderCellBorder();
                cellProp.AppendChild(borders);
            }
        }

        /// <summary>
        /// Add border values to the cell properties
        /// </summary>
        /// <param name="cellProp"></param>
        /// <param name="cell"></param>
        /// <param name="isInAlternateRow"></param>
        private static void AddMargin(this TableCellProperties cellProp, Cell cell, bool isInAlternateRow)
        {
            if (isInAlternateRow && cell.AlternateRowCellConfiguration != null && cell.AlternateRowCellConfiguration.Margin != null)
            {
                cellProp.AppendChild(new TableCellMargin()
                {
                    LeftMargin = new LeftMargin() { Width = cell.AlternateRowCellConfiguration.Margin.Left.ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa },
                    TopMargin = new TopMargin() { Width = cell.AlternateRowCellConfiguration.Margin.Top.ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa },
                    RightMargin = new RightMargin() { Width = cell.AlternateRowCellConfiguration.Margin.Right.ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa },
                    BottomMargin = new BottomMargin() { Width = cell.AlternateRowCellConfiguration.Margin.Bottom.ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa }
                });
            }
            else if (cell.Margin != null)
            {
                cellProp.AppendChild(new TableCellMargin()
                {
                    LeftMargin = new LeftMargin() { Width = cell.Margin.Left.ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa },
                    TopMargin = new TopMargin() { Width = cell.Margin.Top.ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa },
                    RightMargin = new RightMargin() { Width = cell.Margin.Right.ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa },
                    BottomMargin = new BottomMargin() { Width = cell.Margin.Bottom.ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa }
                });
            }
        }
    }
}
