using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Platform.Word.Extensions;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine
{
    public static class ForEachPageExtensions
    {
        public static void Render(this ForEachPage forEach, OpenXmlElement wdDoc, ContextModel context, MainDocumentPart mainDocumentPart, OpenXMLSDK.Word.ReportEngine.Models.Document document)
        {
            context.ReplaceItem(forEach);

            if (!string.IsNullOrEmpty(forEach.DataSourceKey))
            {
                if (context.ExistItem<DataSourceModel>(forEach.DataSourceKey))
                {
                    var datasource = context.GetItem<DataSourceModel>(forEach.DataSourceKey);

                    if (datasource != null && datasource.Items.Count > 0)
                    {
                        foreach (var item in datasource.Items)
                        {
                            var newPage = forEach.Clone();

                            bool addPageBreak = (document.Pages.IndexOf(newPage) < document.Pages.Count - 1);

                            // doc inherit margin from page
                            if (document.Margin == null && newPage.Margin != null)
                                document.Margin = newPage.Margin;
                            // newPage inherit margin from doc
                            else if (document.Margin != null && newPage.Margin == null)
                                newPage.Margin = document.Margin;

                            newPage.Clone().Render(wdDoc, item, mainDocumentPart, addPageBreak);
                        }
                    }
                }
            }
        }
    }
}
