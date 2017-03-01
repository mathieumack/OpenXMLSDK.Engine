using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine
{
    public static class ParagraphExtensions
    {
        public static OpenXmlElement Render(this Paragraph paragraph, OpenXmlElement parent, ContextModel context)
        {
            context.ReplaceItem(paragraph);

            var openXmlPar = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
            parent.Append(openXmlPar);

            return openXmlPar;            
        }
    }
}
