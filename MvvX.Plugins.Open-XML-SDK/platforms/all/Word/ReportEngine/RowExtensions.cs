using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.Word.ReportEngine.BatchModels;
using OpenXMLSDK.Engine.Word.ReportEngine.Models;

namespace OpenXMLSDK.Engine.Word.ReportEngine
{
    public static class RowExtensions
    {
        public static TableRow Render(this Row row, OpenXmlElement parent, ContextModel context, IList<ContextModel> cellsContext, Cell cellModel, OpenXmlPart documentPart, bool isHeader, IFormatProvider formatProvider)
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

            foreach (var cellContext in cellsContext)
            {
                var cell = cellModel.Clone();
                cell.InheritFromParent(row);
                wordRow.AppendChild(cell.Render(wordRow, cellContext, documentPart, formatProvider));
            }

            return wordRow;
        }

        public static TableRow Render(this Row row, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart, bool isHeader, IFormatProvider formatProvider)
        {
            if (!string.IsNullOrWhiteSpace(row.ShowKey) && context.ExistItem<BooleanModel>(row.ShowKey) && !context.GetItem<BooleanModel>(row.ShowKey).Value)
                return null;

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
                wordRowProperties.AppendChild(new TableRowHeight() { Val = UInt32Value.FromUInt32((uint)row.RowHeight.Value)});
            }

            if (row.CantSplit)
            {
                wordRowProperties.AppendChild(new CantSplit());
            }

            foreach (var cell in row.Cells)
            {
                cell.InheritFromParent(row);
                wordRow.AppendChild(cell.Render(wordRow, context, documentPart, formatProvider));
            }

            return wordRow;
        }
    }
}
