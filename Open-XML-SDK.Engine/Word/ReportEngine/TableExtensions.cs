using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.Platform.Word.Extensions;
using OpenXMLSDK.Engine.Word.ReportEngine.BatchModels;

namespace OpenXMLSDK.Engine.Word.ReportEngine
{
    public static class TableExtensions
    {
        /// <summary>
        /// Render a table element
        /// </summary>
        /// <param name="table"></param>
        /// <param name="parent"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        public static Table Render(this OpenXMLSDK.Engine.Word.ReportEngine.Models.Table table, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart)
        {
            context.ReplaceItem(table);

            var createdTable = CreateTable(table, context, documentPart);
            var wordTable = createdTable.Item1;
            var tableLook = createdTable.Item2;

            // Before rows :
            ManageBeforeAfterRows(table, table.BeforeRows, wordTable, context, documentPart);

            // add content rows
            if (!string.IsNullOrEmpty(table.DataSourceKey))
            {
                if (context.ExistItem<DataSourceModel>(table.DataSourceKey))
                {
                    var datasource = context.GetItem<DataSourceModel>(table.DataSourceKey);

                    int i = 0;

                    foreach (var item in datasource.Items)
                    {
                        var row = table.RowModel.Clone();
                        row.InheritFromParent(table);

                        if (!string.IsNullOrWhiteSpace(table.AutoContextAddItemsPrefix))
                        {
                            // We add automatic keys :
                            // Is first item
                            item.AddItem("#" + table.AutoContextAddItemsPrefix + "_TableRow_IsFirstItem#", new BooleanModel(i == 0));
                            // Is last item
                            item.AddItem("#" + table.AutoContextAddItemsPrefix + "_TableRow_IsLastItem#", new BooleanModel(i == datasource.Items.Count - 1));
                            // Index of the element (Based on 0, and based on 1)
                            item.AddItem("#" + table.AutoContextAddItemsPrefix + "_TableRow_IndexBaseZero#", new StringModel(i.ToString()));
                            item.AddItem("#" + table.AutoContextAddItemsPrefix + "_TableRow_IndexBaseOne#", new StringModel((i + 1).ToString()));
                        }

                        wordTable.AppendChild(row.Render(wordTable, item, documentPart, false));

                        i++;
                    }
                }
            }
            else
            {
                foreach (var row in table.Rows)
                {
                    row.InheritFromParent(table);
                    wordTable.AppendChild(row.Render(wordTable, context, documentPart, false));
                }
            }

            // After rows :
            ManageBeforeAfterRows(table, table.AfterRows, wordTable, context, documentPart);

            // Footer
            ManageFooterRow(table, wordTable, tableLook, context, documentPart);

            parent.AppendChild(wordTable);
            return wordTable;
        }

        internal static Tuple<Table, TableLook> CreateTable(OpenXMLSDK.Engine.Word.ReportEngine.Models.Table table, ContextModel context, OpenXmlPart documentPart)
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
            wordTableProperties.TableIndentation = new TableIndentation() { Width = table.TableIndentation.Width, Type = TableWidthUnitValues.Dxa };
            wordTable.AppendChild(wordTableProperties);

            if (table.Borders != null)
            {
                TableBorders borders = table.Borders.Render();

                wordTableProperties.AppendChild(borders);
            }

            // add column width definitions
            if (table.ColsWidth != null)
            {
                wordTable.AppendChild(new TableLayout() { Type = TableLayoutValues.Fixed });

                TableGrid tableGrid = new TableGrid();
                foreach (int width in table.ColsWidth)
                {
                    tableGrid.AppendChild(new GridColumn() { Width = width.ToString(CultureInfo.InvariantCulture) });
                }
                wordTable.AppendChild(tableGrid);
            }

            if (table.TableWidth != null)
            {
                wordTable.AppendChild(new TableWidth() { Width = table.TableWidth.Width, Type = table.TableWidth.Type.ToOOxml() });
            }

            // add header row
            if (table.HeaderRow != null)
            {
                wordTable.AppendChild(table.HeaderRow.Render(wordTable, context, documentPart, true));

                tableLook.FirstRow = OnOffValue.FromBoolean(true);
            }

            return new Tuple<Table, TableLook>(wordTable, tableLook);
        }

        internal static void ManageFooterRow(OpenXMLSDK.Engine.Word.ReportEngine.Models.Table table, Table wordTable, TableLook tableLook, ContextModel context, OpenXmlPart documentPart)
        {
            // add footer row
            if (table.FooterRow != null)
            {
                wordTable.AppendChild(table.FooterRow.Render(wordTable, context, documentPart, false));

                tableLook.LastRow = OnOffValue.FromBoolean(true);
            }
        }

        internal static void ManageBeforeAfterRows(OpenXMLSDK.Engine.Word.ReportEngine.Models.Table table, IList<OpenXMLSDK.Engine.Word.ReportEngine.Models.Row> rows, Table wordTable, ContextModel context, OpenXmlPart documentPart)
        {
            // After rows :
            if (rows != null && rows.Count > 0)
            {
                foreach (var row in rows)
                {
                    row.InheritFromParent(table);
                    wordTable.AppendChild(row.Render(wordTable, context, documentPart, false));
                }
            }
        }
    }
}
