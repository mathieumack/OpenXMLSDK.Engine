using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine
{
    public static class RowExtensions
    {
        public static TableRow Render(this Row row, OpenXmlElement parent, ContextModel context, MainDocumentPart mainDocumentPart)
        {
            context.ReplaceItem(row);

            TableRow wordRow = new TableRow();

            foreach (var cell in row.Cells)
            {
                wordRow.AppendChild(cell.Render(wordRow, context, mainDocumentPart));
            }

            return wordRow;
        }
    }
}
