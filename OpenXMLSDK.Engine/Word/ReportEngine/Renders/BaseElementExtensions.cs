using DO = DocumentFormat.OpenXml;
using DOP = DocumentFormat.OpenXml.Packaging;
using System;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.Extensions;
using ReportEngine.Core.Template.Tables;
using ReportEngine.Core.Template.Images;
using ReportEngine.Core.Template.Charts;
using ReportEngine.Core.Template.Text;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    internal static class BaseElementExtensions
    {
        internal static DO.OpenXmlElement Render(this BaseElement element, Document document, DO.OpenXmlElement parent, ContextModel context, DOP.OpenXmlPart documentPart, IFormatProvider formatProvider)
        {
            context.ReplaceItem(element, formatProvider);

            DO.OpenXmlElement createdElement = null;

            if (element.Show)
            {
                createdElement = element.RenderItem(document, parent, context, documentPart, formatProvider);
            }
            return createdElement;
        }

        private static DO.OpenXmlElement RenderItem(this BaseElement element, Document document, DO.OpenXmlElement parent, ContextModel context, DOP.OpenXmlPart documentPart, IFormatProvider formatProvider)
        {
            DO.OpenXmlElement createdElement = null;

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
            else if (element is PieModel)
            {
                (element as PieModel).Render(parent, context, documentPart, formatProvider);
            }
            else if (element is HtmlContent)
            {
                (element as HtmlContent).Render(parent, context, documentPart, formatProvider);
            }

            if (element.ChildElements != null && element.ChildElements.Count > 0)
            {
                for (int i = 0; i < element.ChildElements.Count; i++)
                {
                    var e = element.ChildElements[i];

                    if (e is TemplateModel)
                    {
                        var elements = (e as TemplateModel).ExtractTemplateItems(document);
                        if (i == element.ChildElements.Count - 1)
                            element.ChildElements.AddRange(elements);
                        else
                            element.ChildElements.InsertRange(i + 1, elements);
                    }
                    else
                    {
                        e.InheritsFromParent(element);
                        e.Render(document, createdElement ?? parent, context, documentPart, formatProvider);
                    }
                }
            }

            return createdElement;
        }
    }
}
