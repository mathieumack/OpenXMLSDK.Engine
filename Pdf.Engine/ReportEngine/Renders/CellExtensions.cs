using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Pdf.Engine.ReportEngine.Extensions;
using Pdf.Engine.ReportEngine.Helpers;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.ExtendedModels;
using ReportEngine.Core.Template.Extensions;
using ReportEngine.Core.Template.Tables;
using ReportEngine.Core.Template.Text;
using it = iTextSharp.text;
using itp = iTextSharp.text.pdf;

namespace Pdf.Engine.ReportEngine.Renders
{
    internal static class CellExtensions
    {
        public static itp.PdfPCell Render(this Cell cell,
                                            Table elementTable,
                                            Document document,
                                            itp.PdfWriter writer,
                                            it.Document pdfDocument,
                                            ContextModel context,
                                            EngineContext ctx,
                                            IFormatProvider formatProvider)
        {
            if (ctx.Parents.Any())
                cell.InheritsFromParent(ctx.Parents.Last());
            ctx.Parents.Add(cell);

            if (cell.Borders == null)
                cell.Borders = elementTable.Borders;

            context.ReplaceItem(cell, formatProvider);

            var pdfCell = new itp.PdfPCell();

            ctx.PCellContainers.Add(cell, pdfCell); // Ajout à la liste des conteneurs du moteur
            
            if(!string.IsNullOrWhiteSpace(cell.Shading))
                pdfCell.BackgroundColor = FontHelper.ConverPdfColorToColor(cell.Shading);
            //pdfCell.FixedHeight = Utilities.MillimetersToPoints(cellModel.Height);
            //pdfCell.NoWrap = cellModel.NoWrap;
            
            if(cell.Justification.HasValue)
                pdfCell.HorizontalAlignment = cell.Justification.Value.ToPdfJustification();
            if (cell.VerticalAlignment.HasValue)
                pdfCell.VerticalAlignment = cell.VerticalAlignment.Value.ToPdfVerticalAlignmentValues();

            pdfCell.Colspan = cell.ColSpan;
            // TODO : Manage rowspan (Fusion field)
            //pdfCell.Rowspan = cellModel.RowSpan;

            //Border width
            if (cell.Borders != null && cell.Borders.UseVariableBorders)
            {
                pdfCell.BorderWidthTop = cell.Borders.BorderPositions.HasFlag(BorderPositions.INSIDEHORIZONTAL) ? cell.Borders.BorderWidthTop / 8f : 0f;
                pdfCell.BorderWidthBottom = cell.Borders.BorderPositions.HasFlag(BorderPositions.INSIDEVERTICAL) ? cell.Borders.BorderWidthBottom / 8f : 0f;
                pdfCell.BorderWidthLeft = cell.Borders.BorderPositions.HasFlag(BorderPositions.LEFT) ? cell.Borders.BorderWidthLeft / 8f : 0f;
                pdfCell.BorderWidthRight = cell.Borders.BorderPositions.HasFlag(BorderPositions.RIGHT) ? cell.Borders.BorderWidthRight / 8f : 0f;
            }
            else if (cell.Borders != null)
                pdfCell.BorderWidth = cell.Borders.BorderWidth / 8f;

            //Border color
            if (cell.Borders != null && !string.IsNullOrWhiteSpace(cell.Borders.BorderColor))
                pdfCell.BorderColor = FontHelper.ConverPdfColorToColor(cell.Borders.BorderColor);

            //Padding
            //if (cell.Padding != null)
            //{
            //    pdfCell.PaddingTop = cell.Padding.Top;
            //    pdfCell.PaddingBottom = cell.Padding.Bottom;
            //    pdfCell.PaddingLeft = cell.Padding.Left;
            //    pdfCell.PaddingRight = cell.Padding.Right;
            //}

            // Render children :
            for (int i = 0; i < cell.ChildElements.Count; i++)
            {
                BaseElement renderElement = null;
                var element = cell.ChildElements[i];
                if (element is Paragraph)
                    renderElement = element;
                else
                {
                    var paragrah = new Paragraph();
                    if (cell.Justification.HasValue)
                        paragrah.Justification = cell.Justification.Value;

                    int j = i;
                    while (j < cell.ChildElements.Count && !(cell.ChildElements[j] is Paragraph))
                    {
                        paragrah.ChildElements.Add(cell.ChildElements[j]);
                        j++;
                    }
                    i = j - 1;
                    renderElement = paragrah;
                }
                renderElement.Render(document, writer, pdfDocument, context, ctx, formatProvider);
            }

            //Add the cell to che collection

            ctx.Parents.RemoveAt(ctx.Parents.Count - 1);
            return pdfCell;
        }
    }
}
