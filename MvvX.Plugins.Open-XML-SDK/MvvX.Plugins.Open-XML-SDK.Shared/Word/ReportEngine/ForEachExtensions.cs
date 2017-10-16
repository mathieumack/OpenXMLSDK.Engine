using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine
{
    public static class ForEachExtensions
    {
        public static void Render(this ForEach forEach, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart)
        {
            context.ReplaceItem(forEach);

            if (!string.IsNullOrEmpty(forEach.DataSourceKey))
            {
                if (context.ExistItem<DataSourceModel>(forEach.DataSourceKey))
                {
                    var datasource = context.GetItem<DataSourceModel>(forEach.DataSourceKey);

                    if (datasource != null && datasource.Items.Count > 0)
                    {
                        int i = 0;
                        string oldKey = null;
                        foreach (var item in datasource.Items)
                        {
                            if (!string.IsNullOrWhiteSpace(oldKey))
                                item.Data.Remove(oldKey);

                            if (!string.IsNullOrWhiteSpace(forEach.AutoContextAddItemsPrefix))
                            {
                                if (i == 0)
                                    item.AddItem("", new BooleanModel(true));
                                if (i == datasource.Items.Count -1)
                                    item.AddItem("", new BooleanModel(true));
                            }
                            foreach (var template in forEach.ItemTemplate)
                            {
                                template.Clone().Render(parent, item, documentPart);
                            }
                            i++;
                        }
                    }
                }
            }
        }
    }
}