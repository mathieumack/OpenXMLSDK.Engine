using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine
{
    public static class PageExtensions
    {
        public static void Render(this Page page, OpenXmlElement wdDoc, ContextModel context, MainDocumentPart mainDocumentPart)
        {
            ((BaseElement)page).Render(wdDoc, context, mainDocumentPart);
        }
    }
}
