using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine
{
    public static class LabelExtension
    {
        public static OpenXmlElement Render(this Label label, OpenXmlElement parent, ContextModel context)
        {
            context.ReplaceItem(label);

            var run = new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(label.Text));
            var runProperty = new DocumentFormat.OpenXml.Wordprocessing.RunProperties()
            {
                RunFonts = new DocumentFormat.OpenXml.Wordprocessing.RunFonts() { Ascii = label.FontName },
                FontSize = new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = label.FontSize },
                Color = new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = label.FontColor }
            };
            run.RunProperties = runProperty;
            parent.Append( run );
            return run;
        }
    }
}
