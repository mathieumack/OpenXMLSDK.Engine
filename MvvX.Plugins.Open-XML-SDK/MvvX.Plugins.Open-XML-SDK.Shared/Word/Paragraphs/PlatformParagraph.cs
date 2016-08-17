using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.Paragraphs;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Paragraphs
{
    public class PlatformParagraph : PlatformOpenXmlElement, IParagraph
    {
        private readonly Paragraph paragraph;

        public PlatformParagraph()
            : this(new Paragraph())
        {
        }

        public PlatformParagraph(Paragraph paragraph)
            : base(paragraph)
        {
            this.paragraph = paragraph;
        }
    }
}
