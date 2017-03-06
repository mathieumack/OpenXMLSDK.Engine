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

            foreach (var element in cell.ChildElements)
            {
                element.Render(wordCell, context, documentPart);
            }

            return wordCell;
        }
    }
}
