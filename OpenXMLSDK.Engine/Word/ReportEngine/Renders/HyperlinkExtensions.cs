using System;
using DO = DocumentFormat.OpenXml;
using DOP = DocumentFormat.OpenXml.Packaging;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template.Text;
using ReportEngine.Core.Template.Extensions;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    public static class HyperlinkExtensions
    {
        /// <summary>
        /// Render a hyperlink.
        /// </summary>
        /// <param name="hyperlink"></param>
        /// <param name="parent"></param>
        /// <param name="context"></param>
        /// <param name="documentPart"></param>
        /// <param name="formatProvider"></param>
        /// <returns></returns>
        public static DO.OpenXmlElement Render(this Hyperlink hyperlink,
                                                    DO.OpenXmlElement parent,
                                                    ContextModel context,
                                                    DOP.OpenXmlPart documentPart,
                                                    IFormatProvider formatProvider)
        {
            context.ReplaceItem(hyperlink, formatProvider);

            if (!hyperlink.Show)
                return null;

            var fieldCodeXmlelement = new DocumentFormat.OpenXml.Wordprocessing.Hyperlink();

            if(!string.IsNullOrWhiteSpace(hyperlink.Anchor))
                fieldCodeXmlelement.Anchor = hyperlink.Anchor;
            else if(!string.IsNullOrWhiteSpace(hyperlink.WebSiteUri))
            {
                var hyperlinkPart = documentPart.AddHyperlinkRelationship(new Uri(hyperlink.WebSiteUri), true);
                fieldCodeXmlelement.Id = hyperlinkPart.Id;
            }

            parent.AppendChild(fieldCodeXmlelement);

            hyperlink.Text.Render(fieldCodeXmlelement, context, documentPart, formatProvider);

            return fieldCodeXmlelement;
        }
    }
}
