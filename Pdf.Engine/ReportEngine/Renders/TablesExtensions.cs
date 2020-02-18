using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Pdf.Engine.ReportEngine.Extensions;
using Pdf.Engine.ReportEngine.Helpers;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.Extensions;
using ReportEngine.Core.Template.Tables;
using ReportEngine.Core.Template.Text;
using it = iTextSharp.text;
using itp = iTextSharp.text.pdf;

namespace Pdf.Engine.ReportEngine.Renders
{
    internal static class TablesExtensions
    {
        /// <summary>
        /// Déclanche le rendu de l'élément
        /// </summary>
        /// <param name="table"></param>
        public static void Render(this Table table,
                                        Document document,
                                        itp.PdfWriter writer,
                                        it.Document pdfDocument,
                                        ContextModel context,
                                        EngineContext ctx,
                                        IFormatProvider formatProvider)
        {
            context.ReplaceItem(table, formatProvider);

            if (!table.Show)
                return;

            table.InheritsFromParent(ctx.Parents.Last());
            ctx.Parents.Add(table);

            int numberOfColumns = table.ColsWidth != null ? table.ColsWidth.Length : 
                                            (table.RowModel != null ? table.RowModel.Cells.Count :
                                                table.Rows[0].Cells.Count);

            itp.PdfPTable pdfTable = new itp.PdfPTable(numberOfColumns);

            pdfTable.KeepTogether = true;

            var columnWidths = new float[] { 50, 50 };

            switch (table.TableWidth.Type)
            {
                case TableWidthUnitValues.Dxa:
                    pdfTable.TotalWidth = table.TableWidth.Width / 2f;
                    if (table.ColsWidth != null)
                    {
                        columnWidths = table.ColsWidth.Select(e => e / 2f).ToArray();
                        pdfTable.SetWidths(columnWidths);
                    }
                    else
                    {
                        // TODO : Calculate col width :
                    }
                    break;
                case TableWidthUnitValues.Pct:
                    //pdfTable.WidthPercentage = table.TableWidth.Width / 50f;
                    pdfTable.TotalWidth = pdfDocument.Right - pdfDocument.LeftMargin;
                    if (table.ColsWidth != null)
                    {
                        columnWidths = table.ColsWidth.Select(e => (e / 50f) / 100).ToArray();
                        pdfTable.SetWidthPercentage(columnWidths, (document.Pages[0] as Page).PageSize.ToRectangle());
                    }
                    else
                    {
                        // TODO : Calculate col width :
                    }
                    break;
            }

            pdfTable.LockedWidth = true;

            var pdfRows = new List<itp.PdfPRow>();

            // Manage header row
            if(table.HeaderRow != null)
            {
                table.HeaderRow.InheritsFromParent(table);
                var pdfRow = table.HeaderRow.Render(table, columnWidths, pdfTable, writer, pdfDocument, document, context, ctx, formatProvider);
                if(pdfRow != null)
                    pdfRows.Add(pdfRow);
            }

            // On regarde si la table est liée à une DataSource et contient donc une RowModel :
            if (!string.IsNullOrWhiteSpace(table.DataSourceKey))
            {
                if (context.ExistItem<DataSourceModel>(table.DataSourceKey))
                {
                    // On va alors transformer le RowModel en Row et l'insérer dans notre table :
                    var datasource = context.GetItem<DataSourceModel>(table.DataSourceKey);
                    for (int i = 0; i < datasource.Items.Count; i++)
                    {
                        var row = table.RowModel.Clone();
                        row.InheritsFromParent(table);

                        if (!string.IsNullOrWhiteSpace(table.AutoContextAddItemsPrefix))
                        {
                            // We add automatic keys :
                            // Is first item
                            var item = datasource.Items[i];
                            item.AddItem("#" + table.AutoContextAddItemsPrefix + "_TableRow_IsFirstItem#", new BooleanModel(i == 0));
                            item.AddItem("#" + table.AutoContextAddItemsPrefix + "_TableRow_IsNotFirstItem#", new BooleanModel(i > 0));
                            // Is last item
                            item.AddItem("#" + table.AutoContextAddItemsPrefix + "_TableRow_IsLastItem#", new BooleanModel(i == datasource.Items.Count - 1));
                            // Index of the element (Based on 0, and based on 1)
                            item.AddItem("#" + table.AutoContextAddItemsPrefix + "_TableRow_IndexBaseZero#", new StringModel(i.ToString()));
                            item.AddItem("#" + table.AutoContextAddItemsPrefix + "_TableRow_IndexBaseOne#", new StringModel((i + 1).ToString()));
                            item.AddItem("#" + table.AutoContextAddItemsPrefix + "_TableRow_IsOdd#", new BooleanModel(i % 2 == 1));
                            item.AddItem("#" + table.AutoContextAddItemsPrefix + "_TableRow_IsEven#", new BooleanModel(i % 2 == 0));
                        }

                        var pdfRow = row.Render(table, columnWidths, pdfTable, writer, pdfDocument, document, datasource.Items[i], ctx, formatProvider);
                        pdfRows.Add(pdfRow);
                    }
                }
            }
            else
            {
                int i = 0;
                foreach (var row in table.Rows)
                {
                    row.InheritsFromParent(table);
                    var pdfRow = row.Render(table, columnWidths, pdfTable, writer, pdfDocument, document, context, ctx, formatProvider);
                    if (pdfRow != null)
                        pdfRows.Add(pdfRow);
                    i++;
                }
            }

            // Manage header row
            if (table.FooterRow != null)
            {
                table.FooterRow.InheritsFromParent(table);
                var pdfRow = table.FooterRow.Render(table, columnWidths, pdfTable, writer, pdfDocument, document, context, ctx, formatProvider);
                if (pdfRow != null)
                    pdfRows.Add(pdfRow);
            }

            table.AddToParentContainer(ctx, pdfTable);

            // Add all rows :
            //pdfTable.Rows.AddRange(pdfRows);

            // TODO : Manage footer row


            ctx.Parents.RemoveAt(ctx.Parents.Count - 1);
        }
    }
}
