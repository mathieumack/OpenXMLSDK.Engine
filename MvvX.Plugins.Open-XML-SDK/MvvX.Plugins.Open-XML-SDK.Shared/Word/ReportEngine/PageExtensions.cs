using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine
{
    public static class PageExtensions
    {
        public static void Render(this Page page, WordprocessingDocument wdDoc, ContextModel context)
        {
            ((BaseElement)page).Render(wdDoc, context);
        }
    }
}
