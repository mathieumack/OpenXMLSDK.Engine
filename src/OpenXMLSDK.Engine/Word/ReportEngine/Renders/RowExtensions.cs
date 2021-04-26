using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.ReportEngine.DataContext;
using OpenXMLSDK.Engine.Word.ReportEngine.Models;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    public static class RowExtensions
    {
        public static TableRow Render(this Row row, Models.Document document, ContextModel context, IList<ContextModel> cellsContext, Cell cellModel, OpenXmlPart documentPart, bool isHeader, bool isAlternateRow, string headerColor, IFormatProvider formatProvider)
        {
            context.ReplaceItem(row, formatProvider);

            TableRow wordRow = new TableRow();

            TableRowProperties wordRowProperties = new TableRowProperties();
            if (isHeader)
            {
                wordRowProperties.AppendChild(new TableHeader() { Val = OnOffOnlyValues.On });
            }
            wordRow.AppendChild(wordRowProperties);

            if (row.RowHeight.HasValue)
            {
                wordRowProperties.AppendChild(new TableRowHeight() { Val = UInt32Value.FromUInt32((uint)row.RowHeight.Value) });
            }

            if (row.CantSplit)
            {
                wordRowProperties.AppendChild(new CantSplit());
            }

            int i = 0;
            foreach (var cellContext in cellsContext)
            {
                var cell = cellModel.Clone();
                // Change shading for row Header
                if (i == 0 && isHeader)
                    cell.Shading = headerColor;
                cell.InheritFromParent(row);
                wordRow.AppendChild(cell.Render(document, wordRow, cellContext, documentPart, isAlternateRow, formatProvider));

                i++;
            }

            return wordRow;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="document"></param>
        /// <param name="parent"></param>
        /// <param name="context"></param>
        /// <param name="documentPart"></param>
        /// <param name="isHeader"></param>
        /// <param name="isAlternateRow">true for even lines</param>
        /// <param name="formatProvider"></param>
        /// <returns></returns>
        public static TableRow Render(this Row row, Models.Document document, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart, bool isHeader, bool isAlternateRow, IFormatProvider formatProvider)
        {
            if (context.TryGetItem(row.ShowKey, out BooleanModel showKey) && !showKey.Value)
            {
                return null;
            }

            context.ReplaceItem(row, formatProvider);

            TableRow wordRow = new TableRow();

            TableRowProperties wordRowProperties = new TableRowProperties();
            if (isHeader)
            {
                wordRowProperties.AppendChild(new TableHeader() { Val = OnOffOnlyValues.On });
            }
            wordRow.AppendChild(wordRowProperties);

            if (row.RowHeight.HasValue)
            {
                wordRowProperties.AppendChild(new TableRowHeight() { Val = UInt32Value.FromUInt32((uint)row.RowHeight.Value) });
            }

            if (row.CantSplit)
            {
                wordRowProperties.AppendChild(new CantSplit());
            }

            foreach (var cell in row.Cells)
            {
                cell.InheritFromParent(row);
                wordRow.AppendChild(cell.Render(document, wordRow, context, documentPart, isAlternateRow, formatProvider));
            }

            return wordRow;
        }
    }
}
