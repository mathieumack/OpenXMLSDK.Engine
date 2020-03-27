using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.ReportEngine.DataContext;
using OpenXMLSDK.Engine.Word.ReportEngine.Models;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    public static class TabulationExtensions
    {
        /// <summary>
        /// Render
        /// </summary>
        /// <param name="tabulation"></param>
        /// <param name="parent"></param>
        /// <param name="context"></param>
        /// <param name="documentPart"></param>
        /// <param name="formatProvider"></param>
        /// <returns></returns>
        public static Tabs Render(this Tabulation tabulation,
                                        OpenXmlElement parent,
                                        ContextModel context,
                                        OpenXmlPart documentPart,
                                        IFormatProvider formatProvider)
        {
            if (!(parent is DocumentFormat.OpenXml.Wordprocessing.Paragraph))
                return null;

            var paragraph = parent as DocumentFormat.OpenXml.Wordprocessing.Paragraph;

            Tabs tabs = new Tabs();
            tabs.AppendChild(new TabStop() { Val = (TabStopValues)tabulation.Alignment, Leader = (DocumentFormat.OpenXml.Wordprocessing.TabStopLeaderCharValues)tabulation.Leader, Position = tabulation.TabStopPosition });

            paragraph.ParagraphProperties.Append(tabs);

            tabulation.Text.IsTabulation = true;
            tabulation.Text.Render(parent, context, documentPart, formatProvider);

            return tabs;
        }
    }
}
