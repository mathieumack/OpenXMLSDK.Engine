using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine
{
    public static class BaseElementExtensions
    {
        public static OpenXmlElement Render(this BaseElement element, OpenXmlElement parent, ContextModel context, MainDocumentPart mainDocumentPart)
        {
            context.ReplaceItem(element);
            OpenXmlElement createdElement = null;

            if (element.Show)
            {
                if (element is Label)
                {
                    createdElement = (element as Label).Render(parent, context);
                }
                else if (element is Paragraph)
                {
                    createdElement = (element as Paragraph).Render(parent, context);
                }
                else if (element is Image)
                {
                    createdElement = (element as Image).Render(parent, context, mainDocumentPart);
                }

                if (element.ChildElements != null && element.ChildElements.Count > 0)
                {
                    foreach (var e in element.ChildElements)
                    {
                        e.InheritFromParent(element);
                        e.Render(createdElement ?? parent, context, mainDocumentPart);
                    }
                }


            }
            return createdElement;
        }
    }
}
