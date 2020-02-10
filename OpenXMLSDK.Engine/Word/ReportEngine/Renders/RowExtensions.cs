using System;
using System.Collections.Generic;
using DO = DocumentFormat.OpenXml;
using DOP = DocumentFormat.OpenXml.Packaging;
using DOW = DocumentFormat.OpenXml.Wordprocessing;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template.Extensions;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.Tables;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    public static class RowExtensions
    {
        public static DOW.TableRow Render(this Row row, Document document, DO.OpenXmlElement parent, ContextModel context, IList<ContextModel> cellsContext, Cell cellModel, DOP.OpenXmlPart documentPart, bool isHeader, bool isAlternateRow, IFormatProvider formatProvider)
        {
            context.ReplaceItem(row, formatProvider);

            var wordRow = new DOW.TableRow();

            var wordRowProperties = new DOW.TableRowProperties();
            if (isHeader)
            {
                wordRowProperties.AppendChild(new DOW.TableHeader() { Val = DOW.OnOffOnlyValues.On });
            }
            wordRow.AppendChild(wordRowProperties);

            if (row.RowHeight.HasValue)
            {
                wordRowProperties.AppendChild(new DOW.TableRowHeight() { Val = DO.UInt32Value.FromUInt32((uint)row.RowHeight.Value) });
            }

            if (row.CantSplit)
            {
                wordRowProperties.AppendChild(new DOW.CantSplit());
            }

            foreach (var cellContext in cellsContext)
            {
                var cell = cellModel.Clone();
                cell.InheritsFromParent(row);
                wordRow.AppendChild(cell.Render(document, wordRow, cellContext, documentPart, isAlternateRow, formatProvider));
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
        public static DOW.TableRow Render(this Row row, Document document, DO.OpenXmlElement parent, ContextModel context, DOP.OpenXmlPart documentPart, bool isHeader, bool isAlternateRow, IFormatProvider formatProvider)
        {
            if (!string.IsNullOrWhiteSpace(row.ShowKey) && context.ExistItem<BooleanModel>(row.ShowKey) && !context.GetItem<BooleanModel>(row.ShowKey).Value)
                return null;

            context.ReplaceItem(row, formatProvider);

            var wordRow = new DOW.TableRow();

            var wordRowProperties = new DOW.TableRowProperties();
            if (isHeader)
            {                
                wordRowProperties.AppendChild(new DOW.TableHeader() { Val = DOW.OnOffOnlyValues.On });
            }
            wordRow.AppendChild(wordRowProperties);

            if (row.RowHeight.HasValue)
            {
                wordRowProperties.AppendChild(new DOW.TableRowHeight() { Val = DO.UInt32Value.FromUInt32((uint)row.RowHeight.Value)});
            }

            if (row.CantSplit)
            {
                wordRowProperties.AppendChild(new DOW.CantSplit());
            }

            foreach (var cell in row.Cells)
            {
                cell.InheritsFromParent(row);
                wordRow.AppendChild(cell.Render(document, wordRow, context, documentPart, isAlternateRow, formatProvider));
            }

            return wordRow;
        }
    }
}
