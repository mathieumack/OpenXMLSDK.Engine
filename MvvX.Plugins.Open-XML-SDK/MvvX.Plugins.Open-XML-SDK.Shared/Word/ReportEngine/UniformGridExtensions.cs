using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Platform.Word.Extensions;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine
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
        public static DocumentFormat.OpenXml.Wordprocessing.Table Render(this OpenXMLSDK.Word.ReportEngine.Models.UniformGrid uniformGrid, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart)
        {
            context.ReplaceItem(uniformGrid);

            if (!string.IsNullOrEmpty(uniformGrid.DataSourceKey) && context.ExistItem<DataSourceModel>(uniformGrid.DataSourceKey))
            {
                var datasource = context.GetItem<DataSourceModel>(uniformGrid.DataSourceKey);

                if (datasource != null && datasource.Items.Count > 0)
                {
                    var createdTable = TableExtensions.CreateTable(uniformGrid, context, documentPart);
                    var wordTable = createdTable.Item1;
                    var tableLook = createdTable.Item2;

                    // Before rows :
                    TableExtensions.ManageBeforeAfterRows(uniformGrid, uniformGrid.BeforeRows, wordTable, context, documentPart);

                    // Table of cells :
                    List<List<ContextModel>> rowsContentContexts = new List<List<ContextModel>>();

                    int i = 0;
                    var lastCellsRow = new List<ContextModel>();

                    // add content rows
                    foreach (var item in datasource.Items)
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

                    var acountAddNullItems = datasource.Items.Count % uniformGrid.ColsWidth.Length;
                    if (acountAddNullItems > 0)
                    {
                        for (int j = 0; j < acountAddNullItems; j++)
                            lastCellsRow.Add(new ContextModel());
                        rowsContentContexts.Add(lastCellsRow);
                    }
                    // Now we create all row :
                    foreach (var rowContentContext in rowsContentContexts)
                    {
                        var row = new Row();
                        row.InheritFromParent(uniformGrid);

                        wordTable.AppendChild(row.Render(wordTable, context, rowContentContext, uniformGrid.CellModel, documentPart, false));

                        i++;
                    }

                    // After rows :
                    TableExtensions.ManageBeforeAfterRows(uniformGrid, uniformGrid.AfterRows, wordTable, context, documentPart);

                    TableExtensions.ManageFooterRow(uniformGrid, wordTable, tableLook, context, documentPart);

                    parent.AppendChild(wordTable);
                    return wordTable;
                }
            }

            return null;
        }
    }
}