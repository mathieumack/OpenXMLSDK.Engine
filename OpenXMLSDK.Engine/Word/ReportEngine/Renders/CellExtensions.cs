using System.Linq;
using ReportEngine.Core.Template;
using System.Globalization;
using OpenXMLSDK.Engine.Platform.Word.Extensions;
using System;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template.Extensions;
using ReportEngine.Core.Template.Text;
using M = ReportEngine.Core.Template;
using ReportEngine.Core.Template.Tables;
using DO = DocumentFormat.OpenXml;
using WP = DocumentFormat.OpenXml.Wordprocessing;
using DOP = DocumentFormat.OpenXml.Packaging;

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
        public static WP.TableCell Render(this Cell cell,
                                            Document document,
                                            DO.OpenXmlElement parent,
                                            ContextModel context,
                                            DOP.OpenXmlPart documentPart,
                                            bool isInAlternateRow,
                                            IFormatProvider formatProvider)
        {
            context.ReplaceItem(cell, formatProvider);

            var wordCell = new WP.TableCell();

            var cellProp = new WP.TableCellProperties();
            wordCell.AppendChild(cellProp);

            if (isInAlternateRow)
                cell.ReplaceAlternateConfiguration();

            cellProp.AddBorders(cell);
            cellProp.AddShading(cell);
            cellProp.AddMargin(cell);
            cellProp.AddJustifications(cell);

            if (cell.CellWidth != null)
                cellProp.AppendChild(new WP.TableCellWidth() { Width = cell.CellWidth.Width, Type = cell.CellWidth.Type.ToOOxml() });

            // manage cell column and row span
            if (cell.ColSpan > 1)
            {
                cellProp.AppendChild(new WP.GridSpan() { Val = cell.ColSpan });
            }
            if (cell.Fusion)
            {
                if (cell.FusionChild)
                {
                    cellProp.AppendChild(new WP.VerticalMerge() { Val = WP.MergedCellValues.Continue });
                }
                else
                {
                    cellProp.AppendChild(new WP.VerticalMerge() { Val = WP.MergedCellValues.Restart });
                }
            }

            if (!cell.Show)
                return wordCell;

            if (cell.ChildElements.Any(x => x is TemplateModel))
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

            if (cell.ChildElements.Any(x => x is Paragraph) || cell.ChildElements.Any(x => x is ForEach))
            {
                foreach (var element in cell.ChildElements)
                {
                    element.InheritsFromParent(cell);
                    if (element is Paragraph
                        || element is ForEach)
                    {
                        element.Render(document, wordCell, context, documentPart, formatProvider);
                    }
                    else
                    {
                        var paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
                        if (cell.Justification.HasValue)
                        {
                            var ppr = new WP.ParagraphProperties();
                            ppr.AppendChild(new WP.Justification() { Val = cell.Justification.Value.ToOOxml() });
                            paragraph.AppendChild(ppr);
                        }
                        wordCell.AppendChild(paragraph);
                        var r = new WP.Run();
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
                    var ppr = new WP.ParagraphProperties();
                    ppr.AppendChild(new WP.Justification() { Val = cell.Justification.Value.ToOOxml() });
                    paragraph.AppendChild(ppr);
                }
                wordCell.AppendChild(paragraph);
                var r = new WP.Run();
                paragraph.AppendChild(r);
                foreach (var element in cell.ChildElements)
                {
                    element.InheritsFromParent(cell);
                    element.Render(document, r, context, documentPart, formatProvider);
                }
            }      

            return wordCell;
        }

        private static void ReplaceAlternateConfiguration(this Cell cell)
        {
            if (cell.AlternateRowCellConfiguration is null)
                return;

            cell.Shading = cell.AlternateRowCellConfiguration.Shading;
            cell.Borders = cell.AlternateRowCellConfiguration.Borders;
            cell.Margin = cell.AlternateRowCellConfiguration.Margin;

            cell.VerticalAlignment = cell.AlternateRowCellConfiguration.VerticalAlignment;
            cell.TextDirection = cell.AlternateRowCellConfiguration.TextDirection;
            cell.CellWidth = cell.AlternateRowCellConfiguration.CellWidth;
        }

        /// <summary>
        /// Add border values to the cell properties
        /// </summary>
        /// <param name="cellProp"></param>
        /// <param name="cell"></param>
        private static void AddJustifications(this WP.TableCellProperties cellProp, Cell cell)
        {
            if (cell.VerticalAlignment.HasValue)
                cellProp.AppendChild(new WP.TableCellVerticalAlignment { Val = cell.VerticalAlignment.Value.ToOOxml() });

            if (cell.TextDirection.HasValue)
                cellProp.AppendChild(new WP.TextDirection { Val = cell.TextDirection.ToOOxml() });
        }

        /// <summary>
        /// Add shading value to the cell properties
        /// </summary>
        /// <param name="cellProp"></param>
        /// <param name="cell"></param>
        /// <param name="isInAlternateRow"></param>
        private static void AddShading(this WP.TableCellProperties cellProp, Cell cell)
        {
            if (!string.IsNullOrEmpty(cell.Shading))
                cellProp.Shading = new WP.Shading() { Fill = cell.Shading };
        }

        /// <summary>
        /// Add border values to the cell properties
        /// </summary>
        /// <param name="cellProp"></param>
        /// <param name="cell"></param>
        /// <param name="isInAlternateRow"></param>
        private static void AddBorders(this WP.TableCellProperties cellProp, Cell cell)
        {
            if (cell.Borders != null)
            {
                var borders = cell.Borders.RenderCellBorder();
                cellProp.AppendChild(borders);
            }
        }

        /// <summary>
        /// Add border values to the cell properties
        /// </summary>
        /// <param name="cellProp"></param>
        /// <param name="cell"></param>
        /// <param name="isInAlternateRow"></param>
        private static void AddMargin(this WP.TableCellProperties cellProp, Cell cell)
        {
            if (cell.Margin != null)
            {
                cellProp.AppendChild(new WP.TableCellMargin()
                {
                    LeftMargin = new WP.LeftMargin() { Width = cell.Margin.Left.ToString(CultureInfo.InvariantCulture), Type = WP.TableWidthUnitValues.Dxa },
                    TopMargin = new WP.TopMargin() { Width = cell.Margin.Top.ToString(CultureInfo.InvariantCulture), Type = WP.TableWidthUnitValues.Dxa },
                    RightMargin = new WP.RightMargin() { Width = cell.Margin.Right.ToString(CultureInfo.InvariantCulture), Type = WP.TableWidthUnitValues.Dxa },
                    BottomMargin = new WP.BottomMargin() { Width = cell.Margin.Bottom.ToString(CultureInfo.InvariantCulture), Type = WP.TableWidthUnitValues.Dxa }
                });
            }
        }
    }
}
