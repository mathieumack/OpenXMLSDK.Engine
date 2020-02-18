using System;
using System.Collections.Generic;
using DO = DocumentFormat.OpenXml;
using DOW = DocumentFormat.OpenXml.Wordprocessing;
using DOP = DocumentFormat.OpenXml.Packaging;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.Tables;
using ReportEngine.Core.Template.Extensions;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    public static class UniformGridExtensions
    {
        /// <summary>
        /// Render a table element
        /// </summary>
        /// <param name="uniformGrid"></param>
        /// <param name="parent"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        public static DOW.Table Render(this UniformGrid uniformGrid, Document document, DO.OpenXmlElement parent, ContextModel context, DOP.OpenXmlPart documentPart, IFormatProvider formatProvider)
        {
            context.ReplaceItem(uniformGrid, formatProvider);

            if (!string.IsNullOrEmpty(uniformGrid.DataSourceKey) && context.ExistItem<DataSourceModel>(uniformGrid.DataSourceKey))
            {
                var dataSource = context.GetItem<DataSourceModel>(uniformGrid.DataSourceKey);

                if (dataSource != null && dataSource.Items.Count > 0)
                {
                    var createdTable = TableExtensions.CreateTable(document, uniformGrid, context, documentPart, formatProvider);
                    var wordTable = createdTable.Item1;
                    var tableLook = createdTable.Item2;

                    // Before rows :
                    TableExtensions.ManageBeforeAfterRows(document, uniformGrid, uniformGrid.BeforeRows, wordTable, context, documentPart, formatProvider);

                    // Table of cells :
                    List<List<ContextModel>> rowsContentContexts = new List<List<ContextModel>>();

                    int i = 0;
                    var lastCellsRow = new List<ContextModel>();

                    // add content rows
                    foreach (var item in dataSource.Items)
                    {
                        if (i > 0 && i % uniformGrid.ColsWidth.Length == 0)
                        {
                            // end of line :
                            rowsContentContexts.Add(lastCellsRow);
                            lastCellsRow = new List<ContextModel>();
                        }
                        lastCellsRow.Add(item);

                        i++;
                    }

                    if (i % uniformGrid.ColsWidth.Length == 0)
                        rowsContentContexts.Add(lastCellsRow);

                    if (dataSource.Items.Count % uniformGrid.ColsWidth.Length > 0)
                    {
                        var acountAddNullItems = uniformGrid.ColsWidth.Length - (dataSource.Items.Count % uniformGrid.ColsWidth.Length);
                        if (acountAddNullItems > 0)
                        {
                            for (int j = 0; j < acountAddNullItems; j++)
                                lastCellsRow.Add(new ContextModel());
                            rowsContentContexts.Add(lastCellsRow);
                        }
                    }

                    // Now we create all row :
                    i = 0;
                    foreach (var rowContentContext in rowsContentContexts)
                    {
                        var row = new Row
                        {
                            CantSplit = uniformGrid.CantSplitRows,
                        };
                        row.InheritsFromParent(uniformGrid);

                        wordTable.AppendChild(row.Render(document, wordTable, context, rowContentContext, uniformGrid.CellModel, documentPart, false, (i % 2 == 1), formatProvider));

                        i++;
                    }

                    // After rows :
                    TableExtensions.ManageBeforeAfterRows(document, uniformGrid, uniformGrid.AfterRows, wordTable, context, documentPart, formatProvider);

                    TableExtensions.ManageFooterRow(document, uniformGrid, wordTable, tableLook, context, documentPart, formatProvider);

                    parent.AppendChild(wordTable);
                    return wordTable;
                }
            }

            return null;
        }
    }
}