using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json;
using OpenXMLSDK.Engine.ReportEngine.DataContext;
using OpenXMLSDK.Engine.Word.ReportEngine.Models;
using OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    internal static class BaseElementExtensions
    {
        internal static T Clone<T>(this T element) where T : BaseElement
        {
            return JsonConvert.DeserializeObject<T>(JsonConvert.SerializeObject(element), new JsonSerializerSettings() { Converters = { new JsonContextConverter() } });
        }

        internal static OpenXmlElement Render(this BaseElement element, Document document, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart, IFormatProvider formatProvider)
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
            if (element is ForEach forEachElement)
            {
                forEachElement.Render(document, parent, context, documentPart, formatProvider);
            }
            else if (element is Label labelElement)
            {
                createdElement = labelElement.Render(parent, context, documentPart, formatProvider);
            }
            else if (element is BookmarkStart bookmarkStartElement)
            {
                createdElement = bookmarkStartElement.Render(parent, context, formatProvider);
            }
            else if (element is BookmarkEnd bookmarkEndElement)
            {
                createdElement = bookmarkEndElement.Render(parent, context, formatProvider);
            }
            else if (element is Hyperlink hyperlinkElement)
            {
                createdElement = hyperlinkElement.Render(parent, context, documentPart, formatProvider);
            }
            else if (element is Paragraph paragraphElement)
            {
                createdElement = paragraphElement.Render(parent, context, formatProvider);
            }
            else if (element is Image imageElement)
            {
                createdElement = imageElement.Render(parent, context, documentPart);
            }
            else if (element is UniformGrid uniformGridElement)
            {
                createdElement = uniformGridElement.Render(document, parent, context, documentPart, formatProvider);
            }
            else if (element is Table tableElement)
            {
                createdElement = tableElement.Render(document, parent, context, documentPart, formatProvider);
            }
            else if (element is TableOfContents tableOfContentsElement)
            {
                tableOfContentsElement.Render(documentPart, context);
            }
            else if (element is BarModel barModelElement)
            {
                barModelElement.Render(parent, context, documentPart, formatProvider);
            }
            else if (element is PieModel pieModelElement)
            {
                pieModelElement.Render(parent, context, documentPart, formatProvider);
            }
            else if (element is LineModel lineModel)
            {
                lineModel.Render(parent, context, documentPart, formatProvider);
            }
            else if (element is HtmlContent htmlContentElement)
            {
                htmlContentElement.Render(parent, context, documentPart, formatProvider);
            }
            else if (element is SimpleField simpleFieldElement)
            {
                simpleFieldElement.Render(parent, context, documentPart, formatProvider);
            }

            if (element.ChildElements != null && element.ChildElements.Count > 0)
            {
                for (int i = 0; i < element.ChildElements.Count; i++)
                {
                    var e = element.ChildElements[i];

                    if (e is TemplateModel templateModelChildElement)
                    {
                        var elements = templateModelChildElement.ExtractTemplateItems(document);
                        if (i == element.ChildElements.Count - 1)
                            element.ChildElements.AddRange(elements);
                        else
                            element.ChildElements.InsertRange(i + 1, elements);
                    }
                    else
                    {
                        e.InheritFromParent(element);
                        e.Render(document, createdElement ?? parent, context, documentPart, formatProvider);
                    }
                }
            }

            return createdElement;
        }
    }
}
