using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenXMLSDK.Word.ReportEngine;
using OpenXMLSDK.Word.ReportEngine.BatchModels;
using OpenXMLSDK.Word.ReportEngine.Models;
using OpenXMLSDK.Word.ReportEngine.Models.Charts;
using Newtonsoft.Json;

namespace OpenXMLSDK.Platform.Word.ReportEngine
{
    public static class BaseElementExtensions
    {
        public static T Clone<T>(this T element) where T : BaseElement
        {
            return JsonConvert.DeserializeObject<T>(JsonConvert.SerializeObject(element), new JsonSerializerSettings() { Converters = { new JsonContextConverter() } });
        }

        public static OpenXmlElement Render(this BaseElement element, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart)
        {
            context.ReplaceItem(element);
            OpenXmlElement createdElement = null;

            if (element.Show)
            {
                // Keep this statement order, because of the UniformGrid inherits from Table
                if (element is ForEach)
                {
                    (element as ForEach).Render(parent, context, documentPart);
                }
                if (element is Label)
                {
                    createdElement = (element as Label).Render(parent, context, documentPart);
                }
                else if (element is Paragraph)
                {
                    createdElement = (element as Paragraph).Render(parent, context);
                }
                else if (element is Image)
                {
                    createdElement = (element as Image).Render(parent, context, documentPart);
                }
                else if (element is UniformGrid)
                {
                    createdElement = (element as UniformGrid).Render(parent, context, documentPart);
                }
                else if (element is Table)
                {
                    createdElement = (element as Table).Render(parent, context, documentPart);
                }
                else if (element is TableOfContents)
                {
                    (element as TableOfContents).Render(documentPart, context);
                }
                else if (element is BarModel)
                {
                    (element as BarModel).Render(parent, context, documentPart);
                }

                if (element.ChildElements != null && element.ChildElements.Count > 0)
                {
                    foreach (var e in element.ChildElements)
                    {
                        e.InheritFromParent(element);
                        e.Render(createdElement ?? parent, context, documentPart);
                    }
                }
            }
            return createdElement;
        }
    }
}
