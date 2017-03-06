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

            if(cell.ColSpan > 1)
            {
                cellProp.AppendChild(new GridSpan() { Val = cell.ColSpan });
            }
            if(cell.Fusion)
            {
                if(cell.FusionChild)
                {
                    cellProp.AppendChild(new VerticalMerge() { Val = MergedCellValues.Continue });
                }
                else
                {
                    cellProp.AppendChild(new VerticalMerge() { Val = MergedCellValues.Restart });
                }
            }

            var paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
            wordCell.AppendChild(paragraph);

            foreach (var element in cell.ChildElements)
            {
                var content = element.Render(paragraph, context, documentPart);
            }

            return wordCell;
        }
    }
}
