using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;
using System.IO;
using System.Linq;
using System.Text;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.ReportEngine
{
    public static class LabelExtension
    {
        public static OpenXmlElement Render(this Label label, OpenXmlElement parent, ContextModel context, OpenXmlPart documentPart)
        {
            context.ReplaceItem(label);

            if (label.isHtml)
            {
                AlternativeFormatImportPart formatImportPart;
                if (documentPart is MainDocumentPart)
                    formatImportPart = (documentPart as MainDocumentPart).AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Xhtml);
                else if (documentPart is HeaderPart)
                    formatImportPart = (documentPart as HeaderPart).AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Xhtml);
                else if (documentPart is FooterPart)
                    formatImportPart = (documentPart as FooterPart).AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Xhtml);
                else
                    return null;

                AltChunk altChunk = new AltChunk();
                altChunk.Id = documentPart.GetIdOfPart(formatImportPart);

                using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(label.Text)))
                {
                    formatImportPart.FeedData(ms);
                }

                OpenXmlElement paragraph = null;
                if (parent is DocumentFormat.OpenXml.Wordprocessing.Paragraph)
                {
                    paragraph = parent;
                }
                else
                {
                    paragraph = parent.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault();
                    
                }
                    
                if (paragraph != null )
                {
                    paragraph.InsertAfterSelf(altChunk);
                }   

                return altChunk;

            }
            else
            {
                var run = new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(label.Text)
                {
                    Space = (SpaceProcessingModeValues)(int)label.SpaceProcessingModeValue
                });
                var runProperty = new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
                if (!string.IsNullOrWhiteSpace(label.FontName))
                    runProperty.RunFonts = new DocumentFormat.OpenXml.Wordprocessing.RunFonts() { Ascii = label.FontName, HighAnsi = label.FontName, EastAsia = label.FontName, ComplexScript = label.FontName };
                if (!string.IsNullOrWhiteSpace(label.FontSize))
                    runProperty.FontSize = new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = label.FontSize };
                if (!string.IsNullOrWhiteSpace(label.FontSize))
                    runProperty.Color = new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = label.FontColor };
                if (!string.IsNullOrWhiteSpace(label.Shading))
                    runProperty.Shading = new DocumentFormat.OpenXml.Wordprocessing.Shading() { Fill = label.Shading };
                if (label.Bold.HasValue)
                    runProperty.Bold = new DocumentFormat.OpenXml.Wordprocessing.Bold() { Val = OnOffValue.FromBoolean(label.Bold.Value) };
                if (label.Italic.HasValue)
                    runProperty.Italic = new DocumentFormat.OpenXml.Wordprocessing.Italic() { Val = OnOffValue.FromBoolean(label.Italic.Value) };
                if (label.IsPageNumber)
                    run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.PageNumber());

                run.RunProperties = runProperty;
                parent.Append(run);

                return run;
            }
        }
    }
}
