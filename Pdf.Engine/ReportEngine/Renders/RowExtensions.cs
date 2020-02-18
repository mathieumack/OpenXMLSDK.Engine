using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Pdf.Engine.ReportEngine.Helpers;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.Extensions;
using ReportEngine.Core.Template.Tables;
using it = iTextSharp.text;
using itp = iTextSharp.text.pdf;

namespace Pdf.Engine.ReportEngine.Renders
{
    internal static class RowExtensions
    {
        /// <summary>
        /// Ajoute une ligne au tableau et renvoie
        /// </summary>
        /// <param name="row"></param>
        /// <param name="elementTable"></param>
        /// <param name="colWidths"></param>
        /// <param name="table"></param>
        /// <param name="writer"></param>
        /// <param name="document"></param>
        /// <param name="ctx"></param>
        /// <param name="ctx"></param>
        /// <returns></returns>
        public static itp.PdfPRow Render(this Row row,
                                                Table elementTable,
                                                float[] colWidths,
                                                itp.PdfPTable table,
                                                itp.PdfWriter writer,
                                                it.Document pdfDocument,
                                                Document document,
                                                ContextModel context,
                                                EngineContext ctx,
                                                IFormatProvider formatProvider)
        {
            context.ReplaceItem(row, formatProvider);
            if (!row.Show)
                return null;

            var cells = new itp.PdfPCell[row.Cells.Count];
            for (int i = 0; i < row.Cells.Count; i++)
            {
                var cellModel = row.Cells[i];

                // Rendu du contenu de la cellule
                var pdfCell = cellModel.Render(elementTable, document, writer, pdfDocument, context, ctx, formatProvider);
                table.AddCell(pdfCell);
                cells[i] = pdfCell;
            }

            //Check the cells number to add a new row if needeed
            var pdfRow = new itp.PdfPRow(cells);
            return pdfRow;
        }
    }
}
