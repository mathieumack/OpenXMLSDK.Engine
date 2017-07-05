using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Platform.Word.Extensions;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;
using Table = DocumentFormat.OpenXml.Drawing.Table;

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
        public static Table Render(this OpenXMLSDK.Word.ReportEngine.Models.UniformGrid uniformGrid, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart)
        {
            context.ReplaceItem(uniformGrid);

            if (!string.IsNullOrEmpty(uniformGrid.DataSourceKey))
            {
                if (context.ExistItem<DataSourceModel>(uniformGrid.DataSourceKey))
                {
                    var datasource = context.GetItem<DataSourceModel>(uniformGrid.DataSourceKey);

                    if (datasource != null && datasource.Items.Count > 0)
                    {
                        Table wordTable = new Table();

                        TableProperties wordTableProperties = new TableProperties();
                        var tableLook = new TableLook()
                        {
                            FirstRow = OnOffValue.FromBoolean(false),
                            LastRow = OnOffValue.FromBoolean(false),
                            FirstColumn = OnOffValue.FromBoolean(false),
                            LastColumn = OnOffValue.FromBoolean(false),
                            NoHorizontalBand = OnOffValue.FromBoolean(false),
                            NoVerticalBand = OnOffValue.FromBoolean(false)
                        };
                        wordTableProperties.AppendChild(tableLook);
                        //The type must be Dxa, if the value is set to Pct : the open xml engine will ignored it !
                        wordTableProperties.TableIndentation = new TableIndentation() { Width = uniformGrid.TableIndentation.Width, Type = TableWidthUnitValues.Dxa };
                        wordTable.AppendChild(wordTableProperties);

                        if (uniformGrid.Borders != null)
                        {
                            TableBorders borders = uniformGrid.Borders.Render();

                            wordTableProperties.AppendChild(borders);
                        }

                        // add column width definitions
                        if (uniformGrid.ColsWidth != null)
                        {
                            wordTable.AppendChild(new TableLayout() { Type = TableLayoutValues.Fixed });

                            TableGrid tableGrid = new TableGrid();
                            foreach (int width in uniformGrid.ColsWidth)
                            {
                                tableGrid.AppendChild(new GridColumn() { Width = width.ToString(CultureInfo.InvariantCulture) });
                            }
                            wordTable.AppendChild(tableGrid);
                        }

                        if (uniformGrid.TableWidth != null)
                        {
                            wordTable.AppendChild(new TableWidth() { Width = uniformGrid.TableWidth.Width, Type = uniformGrid.TableWidth.Type.ToOOxml() });
                        }
                        
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
                        foreach(var rowContentContext in rowsContentContexts)
                        { 
                            var row = new Row();
                            row.InheritFromParent(uniformGrid);

                            wordTable.AppendChild(row.Render(wordTable, context, rowContentContext, uniformGrid.CellModel, documentPart, false));

                            i++;
                        }

                        parent.AppendChild(wordTable);
                        return wordTable;
                    }
                }
            }

            return null;
        }
    }
}