using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json;
using OpenXMLSDK.Engine.Word.ReportEngine.Models;
using OpenXMLSDK.Engine.Word.ReportEngine.BatchModels;
using OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts;
using System;
using OpenXMLSDK.Engine.interfaces.Word.ReportEngine.Models;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    public static class BaseElementExtensions
    {
        public static T Clone<T>(this T element) where T : BaseElement
        {
            return JsonConvert.DeserializeObject<T>(JsonConvert.SerializeObject(element), new JsonSerializerSettings() { Converters = { new JsonContextConverter() } });
        }

        public static OpenXmlElement Render(this BaseElement element, Document document, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart, IFormatProvider formatProvider)
        {
            context.ReplaceItem(element, formatProvider);

            OpenXmlElement createdElement = null;

            if (element.Show)
            {
                createdElement = element.RenderItem(document, parent, context, documentPart, formatProvider);
            }
            return createdElement;
        }

        private static OpenXmlElement RenderItem(this BaseElement element, Document document, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart, IFormatProvider formatProvider)
        {
            OpenXmlElement createdElement = null;

            // Keep this statement order, because of the UniformGrid inherits from Table
            if (element is ForEach)
            {
                (element as ForEach).Render(document, parent, context, documentPart, formatProvider);
            }
            else if (element is Label)
            {
                createdElement = (element as Label).Render(parent, context, documentPart, formatProvider);
            }
            else if (element is BookmarkStart)
            {
                createdElement = (element as BookmarkStart).Render(parent, context, formatProvider);
            }
            else if (element is BookmarkEnd)
            {
                createdElement = (element as BookmarkEnd).Render(parent, context, formatProvider);
            }
            else if (element is Hyperlink)
            {
                createdElement = (element as Hyperlink).Render(parent, context, documentPart, formatProvider);
            }
            else if (element is Paragraph)
            {
                createdElement = (element as Paragraph).Render(parent, context, formatProvider);
            }
            else if (element is Image)
            {
                createdElement = (element as Image).Render(parent, context, documentPart);
            }
            else if (element is UniformGrid)
            {
                createdElement = (element as UniformGrid).Render(document, parent, context, documentPart, formatProvider);
            }
            else if (element is Table)
            {
                createdElement = (element as Table).Render(document, parent, context, documentPart, formatProvider);
            }
            else if (element is TableOfContents)
            {
                (element as TableOfContents).Render(documentPart, context);
            }
            else if (element is BarModel)
            {
                (element as BarModel).Render(parent, context, documentPart, formatProvider);
            }
            else if (element is TemplateModel)
            {
                (element as TemplateModel).Render(document, parent, context, documentPart, formatProvider);
            }
            else if (element is HtmlContent)
            {
                (element as HtmlContent).Render(document, parent, context, documentPart, formatProvider);
            }

            if (element.ChildElements != null && element.ChildElements.Count > 0)
            {
                foreach (var e in element.ChildElements)
                {
                    e.InheritFromParent(element);
                    e.Render(document, createdElement ?? parent, context, documentPart, formatProvider);
                }
            }

            return createdElement;
        }
    }
}
