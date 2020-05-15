using System;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.Platform.Word.Extensions;
using OpenXMLSDK.Engine.ReportEngine.DataContext;
using OpenXMLSDK.Engine.Word.ReportEngine.Models;

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

            if (isInAlternateRow)
                cell.ReplaceAlternateConfiguration();

            cellProp.AddBorders(cell);
            cellProp.AddShading(cell);
            cellProp.AddMargin(cell);
            cellProp.AddJustifications(cell);

            if (cell.CellWidth != null)
                cellProp.AppendChild(new TableCellWidth() { Width = cell.CellWidth.Width, Type = cell.CellWidth.Type.ToOOxml() });

            if (cell.NoWrap)
                cellProp.AppendChild(new NoWrap());

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

            if (!cell.Show)
                return null;

            if (cell.ChildElements.Any(x => x is Models.TemplateModel))
            {
                // Need to replace elements :
                for (int i = 0; i < cell.ChildElements.Count; i++)
                {
                    var e = cell.ChildElements[i];

                    if (e is TemplateModel)
                    {
                        var elements = (e as TemplateModel).ExtractTemplateItems(document);
                        if (i == cell.ChildElements.Count - 1)
                            cell.ChildElements.AddRange(elements);
                        else
                            cell.ChildElements.InsertRange(i + 1, elements);
                    }
                }
            }

            if (cell.ChildElements.Any(x => x is Models.Paragraph) || cell.ChildElements.Any(x => x is ForEach))
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

            return wordCell;
        }

        private static void ReplaceAlternateConfiguration(this Cell cell)
        {
            if (cell.AlternateRowCellConfiguration is null)
                return;

            if(!string.IsNullOrWhiteSpace(cell.AlternateRowCellConfiguration.Shading))
                cell.Shading = cell.AlternateRowCellConfiguration.Shading;
            if (!(cell.AlternateRowCellConfiguration.Borders is null))
                cell.Borders = cell.AlternateRowCellConfiguration.Borders;
            if (!(cell.AlternateRowCellConfiguration.Margin is null))
                cell.Margin = cell.AlternateRowCellConfiguration.Margin;

            if (!(cell.AlternateRowCellConfiguration.VerticalAlignment is null))
                cell.VerticalAlignment = cell.AlternateRowCellConfiguration.VerticalAlignment;
            if (!(cell.AlternateRowCellConfiguration.TextDirection is null))
                cell.TextDirection = cell.AlternateRowCellConfiguration.TextDirection;
            if (!(cell.AlternateRowCellConfiguration.CellWidth is null))
                cell.CellWidth = cell.AlternateRowCellConfiguration.CellWidth;
        }

        /// <summary>
        /// Add border values to the cell properties
        /// </summary>
        /// <param name="cellProp"></param>
        /// <param name="cell"></param>
        private static void AddJustifications(this TableCellProperties cellProp, Cell cell)
        {
            if (cell.VerticalAlignment.HasValue)
                cellProp.AppendChild(new TableCellVerticalAlignment { Val = cell.VerticalAlignment.Value.ToOOxml() });

            if (cell.TextDirection.HasValue)
                cellProp.AppendChild(new TextDirection { Val = cell.TextDirection.ToOOxml() });
        }

        /// <summary>
        /// Add shading value to the cell properties
        /// </summary>
        /// <param name="cellProp"></param>
        /// <param name="cell"></param>
        /// <param name="isInAlternateRow"></param>
        private static void AddShading(this TableCellProperties cellProp, Cell cell)
        {
            if (!string.IsNullOrEmpty(cell.Shading))
                cellProp.Shading = new Shading() { Fill = cell.Shading };
        }

        /// <summary>
        /// Add border values to the cell properties
        /// </summary>
        /// <param name="cellProp"></param>
        /// <param name="cell"></param>
        /// <param name="isInAlternateRow"></param>
        private static void AddBorders(this TableCellProperties cellProp, Cell cell)
        {
            if (cell.Borders != null)
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
        private static void AddMargin(this TableCellProperties cellProp, Cell cell)
        {
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
        }
    }
}
