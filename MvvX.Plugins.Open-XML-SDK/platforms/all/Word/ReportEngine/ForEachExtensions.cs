using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenXMLSDK.Engine.Word.ReportEngine.BatchModels;
using OpenXMLSDK.Engine.Word.ReportEngine.Models;
using System;

namespace OpenXMLSDK.Engine.Word.ReportEngine
{
    public static class ForEachExtensions
    {
        public static void Render(this ForEach forEach, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart, IFormatProvider formatProvider)
        {
            context.ReplaceItem(forEach, formatProvider);

            if (!string.IsNullOrEmpty(forEach.DataSourceKey))
            {
                if (context.ExistItem<DataSourceModel>(forEach.DataSourceKey))
                {
                    var datasource = context.GetItem<DataSourceModel>(forEach.DataSourceKey);

                    if (datasource != null && datasource.Items.Count > 0)
                    {
                        int i = 0;
                        foreach (var item in datasource.Items)
                        {
                            if (!string.IsNullOrWhiteSpace(forEach.AutoContextAddItemsPrefix))
                            {
                                // We add automatic keys :
                                // Is first item
                                item.AddItem("#" + forEach.AutoContextAddItemsPrefix + "_ForEach_IsFirstItem#", new BooleanModel(i == 0));
                                // Is last item
                                item.AddItem("#" + forEach.AutoContextAddItemsPrefix + "_ForEach_IsLastItem#", new BooleanModel(i == datasource.Items.Count - 1));
                                // Index of the element (Based on 0, and based on 1)
                                item.AddItem("#" + forEach.AutoContextAddItemsPrefix + "_ForEach_IndexBaseZero#", new StringModel(i.ToString()));
                                item.AddItem("#" + forEach.AutoContextAddItemsPrefix + "_ForEach_IndexBaseOne#", new StringModel((i+1).ToString()));
                            }

                            foreach (var template in forEach.ItemTemplate)
                            {
                                template.Clone().Render(parent, item, documentPart, formatProvider);
                            }

                            i++;
                        }
                    }
                }
            }
        }
    }
}