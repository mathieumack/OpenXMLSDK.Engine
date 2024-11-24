﻿using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.ReportEngine.DataContext;
using OpenXMLSDK.Engine.Word.Extensions;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    public static class TableExtensions
    {
        /// <summary>
        /// Render a table element
        /// </summary>
        /// <param name="table"></param>
        /// <param name="parent"></param>
        /// <param name="context"></param>
        /// <param name="formatProvider"></param>
        /// <returns></returns>
        public static Table Render(this Models.Table table, Models.Document document, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart, IFormatProvider formatProvider)
        {
            context.ReplaceItem(table, formatProvider);

            var createdTable = CreateTable(document, table, context, documentPart, formatProvider);
            var wordTable = createdTable.Item1;
            var tableLook = createdTable.Item2;

            // Before rows :
            ManageBeforeAfterRows(document, table, table.BeforeRows, wordTable, context, documentPart, formatProvider);

            // add content rows
            if (!string.IsNullOrEmpty(table.DataSourceKey))
            {
                if (context.TryGetItem(table.DataSourceKey, out DataSourceModel dataSource))
                {
                    int i = 0;
                    foreach (var item in dataSource.Items)
                    {
                        var row = table.RowModel.Clone();
                        row.InheritFromParent(table);

                        if (!string.IsNullOrWhiteSpace(table.AutoContextAddItemsPrefix))
                        {
                            // We add automatic keys :
                            // Is first item
                            item.AddItem("#" + table.AutoContextAddItemsPrefix + "_TableRow_IsFirstItem#", new BooleanModel(i == 0));
                            item.AddItem("#" + table.AutoContextAddItemsPrefix + "_TableRow_IsNotFirstItem#", new BooleanModel(i > 0));
                            // Is last item
                            item.AddItem("#" + table.AutoContextAddItemsPrefix + "_TableRow_IsLastItem#", new BooleanModel(i == dataSource.Items.Count - 1));
                            // Index of the element (Based on 0, and based on 1)
                            item.AddItem("#" + table.AutoContextAddItemsPrefix + "_TableRow_IndexBaseZero#", new StringModel(i.ToString()));
                            item.AddItem("#" + table.AutoContextAddItemsPrefix + "_TableRow_IndexBaseOne#", new StringModel((i + 1).ToString()));
                            item.AddItem("#" + table.AutoContextAddItemsPrefix + "_TableRow_IsOdd#", new BooleanModel(i % 2 == 1));
                            item.AddItem("#" + table.AutoContextAddItemsPrefix + "_TableRow_IsEven#", new BooleanModel(i % 2 == 0));
                        }

                        wordTable.AppendChild(row.Render(document, wordTable, item, documentPart, false, (i % 2 == 1), formatProvider));

                        i++;
                    }
                }
            }
            else
            {
                int i = 0;
                foreach (var row in table.Rows)
                {
                    row.InheritFromParent(table);
                    wordTable.AppendChild(row.Render(document, wordTable, context, documentPart, false, (i % 2 == 1), formatProvider));
                    i++;
                }
            }

            // After rows :
            ManageBeforeAfterRows(document, table, table.AfterRows, wordTable, context, documentPart, formatProvider);

            // Footer
            ManageFooterRow(document, table, wordTable, tableLook, context, documentPart, formatProvider);

            parent.AppendChild(wordTable);
            return wordTable;
        }

        internal static Tuple<Table, TableLook> CreateTable(Models.Document document, Models.Table table, ContextModel context, OpenXmlPart documentPart, IFormatProvider formatProvider)
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

            string tableWidth = string.Empty;
            if (table.TableWidth != null)
            {
                tableWidth = table.TableWidth.Width;
                wordTable.AppendChild(new TableWidth() { Width = tableWidth, Type = table.TableWidth.Type.ToOOxml() });

                // If : 
                // - table is a UniformGrid 
                // - no ColsWidth are defined 
                // - tableWidth is an int 
                if (table.GetType() == typeof(Models.UniformGrid)
                    && table.ColsWidth is null
                    && int.TryParse(tableWidth, out int tableWidthValue)
                    )
                {
                    // If Columns are defined on the UniformGrid 
                    if (context.TryGetItem(((Models.UniformGrid)table).ColumnNumberKey, out DoubleModel columnNumberKey)
                        && int.TryParse(columnNumberKey.Value.ToString(), out int columnNumber))
                    {
                        int columnWidth = tableWidthValue / columnNumber;
                        // build ColsWidth before add them
                        table.ColsWidth = new int[columnNumber];
                        for (int i = 0; i < columnNumber; i++)
                        {
                            table.ColsWidth[i] = columnWidth;
                        }

                    }
                }
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

            // add header row
            if (table.HeaderRow != null)
            {
                wordTable.AppendChild(table.HeaderRow.Render(document, wordTable, context, documentPart, true, false, formatProvider));

                tableLook.FirstRow = OnOffValue.FromBoolean(true);
            }

            return new Tuple<Table, TableLook>(wordTable, tableLook);
        }

        internal static void ManageFooterRow(Models.Document document, Models.Table table, Table wordTable, TableLook tableLook, ContextModel context, OpenXmlPart documentPart, IFormatProvider formatProvider)
        {
            // add footer row
            if (table.FooterRow != null)
            {
                wordTable.AppendChild(table.FooterRow.Render(document, wordTable, context, documentPart, false, false, formatProvider));

                tableLook.LastRow = OnOffValue.FromBoolean(true);
            }
        }

        internal static void ManageBeforeAfterRows(Models.Document document, Models.Table table, IList<Models.Row> rows, Table wordTable, ContextModel context, OpenXmlPart documentPart, IFormatProvider formatProvider)
        {
            // After rows :
            if (rows != null && rows.Count > 0)
            {
                foreach (var row in rows)
                {
                    row.InheritFromParent(table);
                    wordTable.AppendChild(row.Render(document, wordTable, context, documentPart, false, false, formatProvider));
                }
            }
        }
    }
}