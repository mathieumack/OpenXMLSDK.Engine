using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.Open_XML_SDK.Core.Word.Paragraphs;

namespace MvvX.Plugins.Open_XML_SDK.Shared.Word.Paragraphs
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
