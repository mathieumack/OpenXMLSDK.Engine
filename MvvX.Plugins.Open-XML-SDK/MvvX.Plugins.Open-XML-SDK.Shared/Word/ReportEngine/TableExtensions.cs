using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;

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
        public static Table Render(this OpenXMLSDK.Word.ReportEngine.Models.Table table, OpenXmlElement parent, ContextModel context, MainDocumentPart mainDocumentPart)
        {
            context.ReplaceItem(table);

            Table wordTable = new Table();

            foreach(var row in table.Rows)
            {
                wordTable.AppendChild(row.Render(wordTable, context, mainDocumentPart));
            }

            return wordTable;
        }
    }
}
