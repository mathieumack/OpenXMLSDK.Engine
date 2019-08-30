using System;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.interfaces.Word.ReportEngine.Models;
using OpenXMLSDK.Engine.Word.ReportEngine.BatchModels;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    public static class HtmlContentExtensions
    {
        /// <summary>
        /// Render a label
        /// </summary>
        /// <param name="label"></param>
        /// <param name="parent"></param>
        /// <param name="context"></param>
        /// <param name="documentPart"></param>
        /// <param name="formatProvider"></param>
        /// <returns></returns>
        public static OpenXmlElement Render(this HtmlContent label, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart, IFormatProvider formatProvider)
        {
            context.ReplaceItem(label, formatProvider);

            AlternativeFormatImportPart formatImportPart;
            if (documentPart is MainDocumentPart)
                formatImportPart = (documentPart as MainDocumentPart).AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Xhtml);
            else if (documentPart is HeaderPart)
                formatImportPart = (documentPart as HeaderPart).AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Xhtml);
            else if (documentPart is FooterPart)
                formatImportPart = (documentPart as FooterPart).AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Xhtml);
            else
                return null;

            return SetHtmlContent(label, parent, documentPart, formatImportPart);
        }

        /// <summary>
        /// Set html content.
        /// </summary>
        /// <param name="label"></param>
        /// <param name="parent"></param>
        /// <param name="documentPart"></param>
        /// <param name="formatImportPart"></param>
        /// <returns></returns>
        private static AltChunk SetHtmlContent(HtmlContent label, OpenXmlElement parent, OpenXmlPart documentPart, AlternativeFormatImportPart formatImportPart)
        {
            AltChunk altChunk = new AltChunk();
            altChunk.Id = documentPart.GetIdOfPart(formatImportPart);

            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(label.Text)))
            {
                formatImportPart.FeedData(ms);
            }

            parent.Append(altChunk);

            return altChunk;
        }
    }
}
