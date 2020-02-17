using System;
using System.Collections.Generic;
using System.Globalization;
using DO = DocumentFormat.OpenXml;
using DOP = DocumentFormat.OpenXml.Packaging;
using DOW = DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.Platform.Word.Extensions;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.Extensions;
using ReportEngine.Core.Template.Tables;
using ReportEngine.Core.DataContext;

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
        public static DOW.Table Render(this Table table, Document document, DO.OpenXmlElement parent, ContextModel context, DOP.OpenXmlPart documentPart, IFormatProvider formatProvider)
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
                if (context.ExistItem<DataSourceModel>(table.DataSourceKey))
                {
                    var datasource = context.GetItem<DataSourceModel>(table.DataSourceKey);

                    int i = 0;

                    foreach (var item in datasource.Items)
                    {
                        var row = table.RowModel.Clone();
                        row.InheritsFromParent(table);

                        if (!string.IsNullOrWhiteSpace(table.AutoContextAddItemsPrefix))
                        {
                            // We add automatic keys :
                            // Is first item
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

                        var rowRender = row.Render(document, wordTable, item, documentPart, false, (i % 2 == 1), formatProvider);
                        wordTable.AppendChild(rowRender);

                        i++;
                    }
                }
            }
            else
            {
                int i = 0;
                foreach (var row in table.Rows)
                {
                    row.InheritsFromParent(table);
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

        internal static Tuple<DOW.Table, DOW.TableLook> CreateTable(Document document, Table table, ContextModel context, DOP.OpenXmlPart documentPart, IFormatProvider formatProvider)
        {
            var wordTable = new DOW.Table();

            var wordTableProperties = new DOW.TableProperties();
            var tableLook = new DOW.TableLook()
            {
                FirstRow = DO.OnOffValue.FromBoolean(false),
                LastRow = DO.OnOffValue.FromBoolean(false),
                FirstColumn = DO.OnOffValue.FromBoolean(false),
                LastColumn = DO.OnOffValue.FromBoolean(false),
                NoHorizontalBand = DO.OnOffValue.FromBoolean(false),
                NoVerticalBand = DO.OnOffValue.FromBoolean(false)
            };
            wordTableProperties.AppendChild(tableLook);
            //The type must be Dxa, if the value is set to Pct : the open xml engine will ignored it !
            wordTableProperties.TableIndentation = new DOW.TableIndentation() { Width = table.TableIndentation.Width, Type = table.TableIndentation.Type.ToOOxml() };
            wordTable.AppendChild(wordTableProperties);

            if (table.Borders != null)
            {
                var borders = table.Borders.Render();

                wordTableProperties.AppendChild(borders);
            }

            // add column width definitions
            if (table.ColsWidth != null)
            {
                wordTable.AppendChild(new DOW.TableLayout() { Type = DOW.TableLayoutValues.Fixed });

                var tableGrid = new DOW.TableGrid();
                foreach (int width in table.ColsWidth)
                {
                    tableGrid.AppendChild(new DOW.GridColumn() { Width = width.ToString(CultureInfo.InvariantCulture) });
                }
                wordTable.AppendChild(tableGrid);
            }

            if (table.TableWidth != null)
            {
                wordTable.AppendChild(new DOW.TableWidth() { Width = table.TableWidth.Width.ToString(), Type = table.TableWidth.Type.ToOOxml() });
            }

            // add header row
            if (table.HeaderRow != null)
            {
                wordTable.AppendChild(table.HeaderRow.Render(document, wordTable, context, documentPart, true, false, formatProvider));

                tableLook.FirstRow = DO.OnOffValue.FromBoolean(true);
            }

            return new Tuple<DOW.Table, DOW.TableLook>(wordTable, tableLook);
        }

        internal static void ManageFooterRow(Document document, Table table, DOW.Table wordTable, DOW.TableLook tableLook, ContextModel context, DOP.OpenXmlPart documentPart, IFormatProvider formatProvider)
        {
            // add footer row
            if (table.FooterRow != null)
            {
                wordTable.AppendChild(table.FooterRow.Render(document, wordTable, context, documentPart, false, false, formatProvider));

                tableLook.LastRow = DO.OnOffValue.FromBoolean(true);
            }
        }

        internal static void ManageBeforeAfterRows(Document document, Table table, IList<Row> rows, DOW.Table wordTable, ContextModel context, DOP.OpenXmlPart documentPart, IFormatProvider formatProvider)
        {
            // After rows :
            if (rows != null && rows.Count > 0)
            {
                foreach (var row in rows)
                {
                    row.InheritsFromParent(table);
                    wordTable.AppendChild(row.Render(document, wordTable, context, documentPart, false, false, formatProvider));
                }
            }
        }
    }
}
