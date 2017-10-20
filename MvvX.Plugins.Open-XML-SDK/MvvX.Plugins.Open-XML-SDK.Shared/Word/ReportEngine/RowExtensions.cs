using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine
{
    public static class RowExtensions
    {
        public static TableRow Render(this Row row, OpenXmlElement parent, ContextModel context, IList<ContextModel> cellsContext, Cell cellModel, OpenXmlPart documentPart, bool isHeader)
        {
            TableRow wordRow = Render(row, context, isHeader);

            foreach (var cellContext in cellsContext)
            {
                var cell = cellModel.Clone();
                cell.InheritFromParent(row);
                wordRow.AppendChild(cell.Render(wordRow, cellContext, documentPart));
            }

            return wordRow;
        }

        public static TableRow Render(this Row row, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart, bool isHeader)
        {
            if (!string.IsNullOrWhiteSpace(row.ShowKey) && context.ExistItem<BooleanModel>(row.ShowKey) && !context.GetItem<BooleanModel>(row.ShowKey).Value)
                return null;

            TableRow wordRow = Render(row, context, isHeader);

            foreach (var cell in row.Cells)
            {
                cell.InheritFromParent(row);
                wordRow.AppendChild(cell.Render(wordRow, context, documentPart));
            }

            return wordRow;
        }


        /// <summary>
        /// Commun Render 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="context"></param>
        /// <param name="isHeader"></param>
        /// <returns></returns>
        /// <remarks>This avoid a code duplication</remarks>
        private static TableRow Render(this Row row, ContextModel context, bool isHeader)
        {
            context.ReplaceItem(row);

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

            return wordRow;
        }
    }
}
