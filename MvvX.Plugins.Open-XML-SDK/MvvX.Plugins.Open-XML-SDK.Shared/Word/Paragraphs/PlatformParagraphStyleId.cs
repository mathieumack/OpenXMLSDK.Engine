using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.Paragraphs;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Paragraphs
{
    public class PlatformParagraphStyleId : PlatformOpenXmlElement, IParagraphStyleId
    {
        private readonly ParagraphStyleId paragraphStyleId;

        public PlatformParagraphStyleId()
            : this(new ParagraphStyleId())
        {
        }

        public PlatformParagraphStyleId(ParagraphStyleId psi)
            : base(psi)
        {
            this.paragraphStyleId = psi;
        }

        public string Val
        {
            get
            {
                return paragraphStyleId.Val?.Value;
            }
            set
            {
                paragraphStyleId.Val = new DocumentFormat.OpenXml.StringValue(value);
            }
        }

        public static PlatformParagraphStyleId New(ParagraphProperties paragraphProperties)
        {
            var paragraphStyleId = CheckDescendantsOrAppendNewOne<ParagraphStyleId>(paragraphProperties);
            return new PlatformParagraphStyleId(paragraphStyleId);
        }
    }
}