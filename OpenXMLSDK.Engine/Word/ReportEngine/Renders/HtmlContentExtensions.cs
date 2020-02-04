using System;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.Extensions;

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

            // encode specifi caracteres :
            label.Text = label.Text.Replace("©", "&#169;")
                                .Replace("ª", "&#170;")
                                .Replace("«", "&#171;")
                                .Replace("¬", "&#172;")
                                .Replace("­", "&#173;")
                                .Replace("®", "&#174;")
                                .Replace("¯", "&#175;")
                                .Replace("°", "&#176;")
                                .Replace("±", "&#177;")
                                .Replace("²", "&#178;")
                                .Replace("³", "&#179;")
                                .Replace("´", "&#180;")
                                .Replace("µ", "&#181;")
                                .Replace("¶", "&#182;")
                                .Replace("·", "&#183;")
                                .Replace("¸", "&#184;")
                                .Replace("¹", "&#185;")
                                .Replace("º", "&#186;")
                                .Replace("»", "&#187;")
                                .Replace("¼", "&#188;")
                                .Replace("½", "&#189;")
                                .Replace("¾", "&#190;")
                                .Replace("¿", "&#191;")
                                .Replace("À", "&#192;")
                                .Replace("Á", "&#193;")
                                .Replace("Â", "&#194;")
                                .Replace("Ã", "&#195;")
                                .Replace("Ä", "&#196;")
                                .Replace("Å", "&#197;")
                                .Replace("Æ", "&#198;")
                                .Replace("Ç", "&#199;")
                                .Replace("È", "&#200;")
                                .Replace("É", "&#201;")
                                .Replace("Ê", "&#202;")
                                .Replace("Ë", "&#203;")
                                .Replace("Ì", "&#204;")
                                .Replace("Í", "&#205;")
                                .Replace("Î", "&#206;")
                                .Replace("Ï", "&#207;")
                                .Replace("Ð", "&#208;")
                                .Replace("Ñ", "&#209;")
                                .Replace("Ò", "&#210;")
                                .Replace("Ó", "&#211;")
                                .Replace("Ô", "&#212;")
                                .Replace("Õ", "&#213;")
                                .Replace("Ö", "&#214;")
                                .Replace("×", "&#215;")
                                .Replace("Ø", "&#216;")
                                .Replace("Ù", "&#217;")
                                .Replace("Ú", "&#218;")
                                .Replace("Û", "&#219;")
                                .Replace("Ü", "&#220;")
                                .Replace("Ý", "&#221;")
                                .Replace("Þ", "&#222;")
                                .Replace("ß", "&#223;")
                                .Replace("à", "&#224;")
                                .Replace("á", "&#225;")
                                .Replace("â", "&#226;")
                                .Replace("ã", "&#227;")
                                .Replace("ä", "&#228;")
                                .Replace("å", "&#229;")
                                .Replace("æ", "&#230;")
                                .Replace("ç", "&#231;")
                                .Replace("è", "&#232;")
                                .Replace("é", "&#233;")
                                .Replace("ê", "&#234;")
                                .Replace("ë", "&#235;")
                                .Replace("ì", "&#236;")
                                .Replace("í", "&#237;")
                                .Replace("î", "&#238;")
                                .Replace("ï", "&#239;")
                                .Replace("ð", "&#240;")
                                .Replace("ñ", "&#241;")
                                .Replace("ò", "&#242;")
                                .Replace("ó", "&#243;")
                                .Replace("ô", "&#244;")
                                .Replace("õ", "&#245;")
                                .Replace("ö", "&#246;")
                                .Replace("÷", "&#247;")
                                .Replace("ø", "&#248;")
                                .Replace("ù", "&#249;")
                                .Replace("ú", "&#250;")
                                .Replace("û", "&#251;")
                                .Replace("ü", "&#252;")
                                .Replace("ý", "&#253;");
            
            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(label.Text)))
            {
                formatImportPart.FeedData(ms);
            }

            parent.Append(altChunk);

            return altChunk;
        }
    }
}
