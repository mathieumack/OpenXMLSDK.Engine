using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.Paragraphs;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Paragraphs
{
    public class PlatformParagraph : PlatformOpenXmlElement, IParagraph
    {
        private readonly Paragraph xmlElement;

        private IParagraphProperties properties;
        /// <summary>
        /// Properties of the paragraph item
        /// </summary>
        public IParagraphProperties ParagraphProperties
        {
            get
            {
                if (properties == null)
                    properties = PlatformParagraphProperties.New(xmlElement);
                return properties;
            }
        }

        public PlatformParagraph()
            : this(new Paragraph())
        {
        }

        public PlatformParagraph(Paragraph paragraph)
            : base(paragraph)
        {
            this.xmlElement = paragraph;
        }
    }
}
