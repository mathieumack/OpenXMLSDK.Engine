using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine
{
    public static class BaseElementExtensions
    {
        public static void Render(this BaseElement element, WordprocessingDocument wdDoc, ContextModel context)
        {
            context.ReplaceItem(element);
            if(element.Show)
            {
                if(element.ChildElements != null && element.ChildElements.Count > 0)
                {
                    foreach(var e in element.ChildElements)
                    {
                        e.Render(wdDoc, context);
                    }
                }

                if(element is Label)
                {
                    (element as Label).Render(wdDoc, context);
                }
            }
        }
    }
}
