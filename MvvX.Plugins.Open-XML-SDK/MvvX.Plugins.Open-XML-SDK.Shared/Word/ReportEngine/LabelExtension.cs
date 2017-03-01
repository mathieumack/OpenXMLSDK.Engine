using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine
{
    public static class LabelExtension
    {
        public static void Render(this Label label, WordprocessingDocument wdDoc, ContextModel context)
        {
            context.ReplaceItem(label);

            wdDoc.MainDocumentPart.Document.Body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(  new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(label.Text))));
        }
    }
}
