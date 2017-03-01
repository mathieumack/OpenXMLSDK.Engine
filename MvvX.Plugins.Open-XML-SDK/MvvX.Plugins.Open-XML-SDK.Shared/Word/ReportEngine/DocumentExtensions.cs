using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine
{
    public static class DocumentExtensions
    {
        public static void Render(this Document document, WordprocessingDocument wdDoc, ContextModel context)
        {

            foreach(var page in document.Pages)
            {
                page.Render(wdDoc, context);
            }
        }
    }
}
