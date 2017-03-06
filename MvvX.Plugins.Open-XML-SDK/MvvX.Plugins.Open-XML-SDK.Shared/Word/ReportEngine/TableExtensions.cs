using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models.Attributes;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine
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
        public static Table Render(this OpenXMLSDK.Word.ReportEngine.Models.Table table, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart)
        {
            context.ReplaceItem(table);

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
            wordTable.AppendChild(wordTableProperties);

            if (table.Borders != null)
            {
                TableBorders borders = table.Borders.Render();

                wordTableProperties.AppendChild(borders);
            }

            // add header row
            if (table.HeaderRow != null)
            {
                wordTable.AppendChild(table.HeaderRow.Render(wordTable, context, documentPart, true));

                tableLook.FirstRow = OnOffValue.FromBoolean(true);
            }

            // add content rows
            foreach (var row in table.Rows)
            {
                wordTable.AppendChild(row.Render(wordTable, context, documentPart, false));
            }

            // add footer row
            if (table.FooterRow != null)
            {
                wordTable.AppendChild(table.FooterRow.Render(wordTable, context, documentPart, false));

                tableLook.LastRow = OnOffValue.FromBoolean(true);
            }
            parent.AppendChild(wordTable);
            return wordTable;
        }
    }
}
