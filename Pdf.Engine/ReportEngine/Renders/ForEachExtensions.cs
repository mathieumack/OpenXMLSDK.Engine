using it = iTextSharp.text;
using itp = iTextSharp.text.pdf;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using System;
using ReportEngine.Core.Template.Extensions;

namespace Pdf.Engine.ReportEngine.Renders
{
    internal static class ForEachExtensions
    {
        public static void Render(this ForEach forEach,
                                        Document document,
                                        itp.PdfWriter writer,
                                        it.Document pdfDocument,
                                        ContextModel context,
                                        EngineContext ctx,
                                        IFormatProvider formatProvider)
        {
            context.ReplaceItem(forEach, formatProvider);
            if (!forEach.Show)
                return;

            if (string.IsNullOrEmpty(forEach.DataSourceKey)
                || !context.ExistItem<DataSourceModel>(forEach.DataSourceKey))
                return;

            var datasource = context.GetItem<DataSourceModel>(forEach.DataSourceKey);
            if (datasource == null || datasource.Items == null)
                return;

            int i = 0;
            foreach (var item in datasource.Items)
            {
                forEach.AddAutoContextAddItemsPrefix(item, i, datasource);

                for (int j = 0; j < forEach.ItemTemplate.Count; j++)
                {
                    var template = forEach.ItemTemplate[j];
                    var clone = template.Clone();

                    // Specific rule for TemplateModels in childs :
                    if (clone is TemplateModel)
                    {
                        var templateElements = (clone as TemplateModel).ExtractTemplateItems(document);
                        foreach (var templateElement in templateElements)
                            templateElement.Render(document, writer, pdfDocument, context, ctx, formatProvider);
                    }
                    else
                        clone.Render(document, writer, pdfDocument, context, ctx, formatProvider);
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
        private static void AddAutoContextAddItemsPrefix(this ForEach forEach, ContextModel item, int i, DataSourceModel datasource)
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
