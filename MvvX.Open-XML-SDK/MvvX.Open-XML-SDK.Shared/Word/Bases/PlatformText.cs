using DocumentFormat.OpenXml.Wordprocessing;

namespace MvvX.Open_XML_SDK.Core.Word.Bases
{
    public class PlatformText : PlatformOpenXmlElement, IText
    {
        private readonly Text text;

        public PlatformText(Text text)
            : base(text)
        {
            this.text = text;
        }
    }
}
