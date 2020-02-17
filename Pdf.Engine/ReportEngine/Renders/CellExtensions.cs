using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
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

            var pdfCell = new itp.PdfPCell();

            ctx.PCellContainers.Add(cell, pdfCell); // Ajout à la liste des conteneurs du moteur
            
            //cell.BackgroundColor = FontHelper.ConverPdfColorToColor(cellModel.BackgroundColor);
            //cell.FixedHeight = Utilities.MillimetersToPoints(cellModel.Height);
            //cell.NoWrap = cellModel.NoWrap;
            //cell.HorizontalAlignment = (int)cellModel.HorizontalAlignment;
            //cell.VerticalAlignment = (int)cellModel.VerticalAlignment;
            pdfCell.Colspan = cell.ColSpan;
            // TODO : Manage rowspan (Fusion field)
            //cell.Rowspan = cellModel.RowSpan;

            //Border width
            //if (cellModel.BorderModel.UseVariableBorders)
            //{
            //    cell.BorderWidthTop = cellModel.BorderModel.BorderWidthTop;
            //    cell.BorderWidthBottom = cellModel.BorderModel.BorderWidthBottom;
            //    cell.BorderWidthLeft = cellModel.BorderModel.BorderWidthLeft;
            //    cell.BorderWidthRight = cellModel.BorderModel.BorderWidthRight;
            //}
            //else
            //    cell.BorderWidth = cellModel.BorderModel.BorderWidth;

            //Border color
            //cell.BorderColor = FontHelper.ConverPdfColorToColor(cellModel.BorderModel.BorderColor);

            //Padding
            //if (cellModel.Padding != null)
            //{
            //    cell.PaddingTop = cellModel.Padding.Top;
            //    cell.PaddingBottom = cellModel.Padding.Bottom;
            //    cell.PaddingLeft = cellModel.Padding.Left;
            //    cell.PaddingRight = cellModel.Padding.Right;
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
