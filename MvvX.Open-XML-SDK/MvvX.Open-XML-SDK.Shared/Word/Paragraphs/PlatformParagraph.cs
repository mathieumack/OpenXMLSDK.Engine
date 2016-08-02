using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word.Bases;

namespace MvvX.Open_XML_SDK.Core.Word.Paragraphs
{
    public class PlatformParagraph : PlatformOpenXmlElement, IParagraph
    {
        private readonly Paragraph paragraph;
        private readonly PlatformOpenXmlElement[] childElements;
        public PlatformParagraph(Paragraph paragraph)
            : base(paragraph)
        {
            this.paragraph = paragraph;
        }

        public PlatformParagraph()
        {

        }

        //public PlatformParagraph(params PlatformOpenXmlElement[] childElements)
        //{
        //    this.childElements = childElements;
        //}

    }
}
