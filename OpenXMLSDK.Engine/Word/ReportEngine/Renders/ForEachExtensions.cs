using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.Extensions;
using System;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    public static class ForEachExtensions
    {
        /// <summary>
        /// Render a ForEach element in the document
        /// </summary>
        /// <param name="forEach"></param>
        /// <param name="document"></param>
        /// <param name="parent"></param>
        /// <param name="context"></param>
        /// <param name="documentPart"></param>
        /// <param name="formatProvider"></param>
        public static void Render(this ForEach forEach, 
                                        Document document, 
                                        OpenXmlElement parent, 
                                        ContextModel context, 
                                        OpenXmlPart documentPart, 
                                        IFormatProvider formatProvider)
        {
            context.ReplaceItem(forEach, formatProvider);

            if (string.IsNullOrEmpty(forEach.DataSourceKey)
                || !context.ExistItem<DataSourceModel>(forEach.DataSourceKey))
                    return;

            var datasource = context.GetItem<DataSourceModel>(forEach.DataSourceKey);
            if (datasource == null || datasource.Items == null)
                return;

            int i = 0;
            foreach (var item in datasource.Items)
            {
                item.AddAutoContextAddItemsPrefix(forEach, i, datasource);

                for (int j = 0; j < forEach.ItemTemplate.Count; j++)
                {
                    var template = forEach.ItemTemplate[j];
                    var clone = template.Clone();

                    // Specific rule for TemplateModels in childs :
                    if (clone is TemplateModel)
                    {
                        var templateElements = (clone as TemplateModel).ExtractTemplateItems(document);
                        foreach (var templateEement in templateElements)
                            templateEement.Render(document, parent, item, documentPart, formatProvider);
                    }
                    else
                        clone.Render(document, parent, item, documentPart, formatProvider);
                }

                i++;
            }
        }

        /// <summary>
        /// Add automatic keys for the foreach item
        /// </summary>
        /// <param name="item"></param>
        /// <param name="forEach"></param>
        /// <param name="i"></param>
        /// <param name="datasource"></param>
        private static void AddAutoContextAddItemsPrefix(this ContextModel item, ForEach forEach, int i, DataSourceModel datasource)
        {
            if (forEach == null)
                return;

            if (!string.IsNullOrWhiteSpace(forEach.AutoContextAddItemsPrefix))
            {
                // We add automatic keys :
                // Is first item
                item.AddItem("#" + forEach.AutoContextAddItemsPrefix + "_ForEach_IsFirstItem#", new BooleanModel(i == 0));
                item.AddItem("#" + forEach.AutoContextAddItemsPrefix + "_ForEach_IsNotFirstItem#", new BooleanModel(i > 0));
                // Is last item
                item.AddItem("#" + forEach.AutoContextAddItemsPrefix + "_ForEach_IsLastItem#", new BooleanModel(i == datasource.Items.Count - 1));
                // Index of the element (Based on 0, and based on 1)
                item.AddItem("#" + forEach.AutoContextAddItemsPrefix + "_ForEach_IndexBaseZero#", new StringModel(i.ToString()));
                item.AddItem("#" + forEach.AutoContextAddItemsPrefix + "_ForEach_IndexBaseOne#", new StringModel((i + 1).ToString()));
                item.AddItem("#" + forEach.AutoContextAddItemsPrefix + "_ForEach_IsOdd#", new BooleanModel(i % 2 == 1));
                item.AddItem("#" + forEach.AutoContextAddItemsPrefix + "_ForEach_IsEven#", new BooleanModel(i % 2 == 0));
            }
        }
    }
}