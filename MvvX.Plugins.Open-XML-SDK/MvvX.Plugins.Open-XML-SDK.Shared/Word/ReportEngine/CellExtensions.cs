using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine
{
    public static class CellExtensions
    {
        public static TableCell Render(this Cell cell, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart)
        {
            context.ReplaceItem(cell);

            TableCell wordCell = new TableCell();

            TableCellProperties cellProp = new TableCellProperties();
            wordCell.AppendChild(cellProp);

            if (cell.Borders != null)
            {
                TableCellBorders borders = cell.Borders.RenderCellBorder();

                cellProp.AppendChild(borders);
            }

            foreach (var element in cell.ChildElements)
            {
                var paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
                wordCell.AppendChild(paragraph);
                var content = element.Render(paragraph, context, documentPart);
            }

            return wordCell;
        }
    }
}
