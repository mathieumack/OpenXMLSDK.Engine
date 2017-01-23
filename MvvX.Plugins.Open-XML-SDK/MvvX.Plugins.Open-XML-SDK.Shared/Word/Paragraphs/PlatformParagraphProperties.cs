using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.Paragraphs;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Paragraphs
{
    public class PlatformParagraphProperties : PlatformOpenXmlElement, IParagraphProperties
    {
        private readonly ParagraphProperties xmlElement;

        public PlatformParagraphProperties()
            : this(new ParagraphProperties())
        {
        }

        public PlatformParagraphProperties(ParagraphProperties paragraph)
            : base(paragraph)
        {
            this.xmlElement = paragraph;
        }

        private INumberingProperties numberingProperties;
        public INumberingProperties NumberingProperties
        {
            get
            {
                if (numberingProperties == null)
                    numberingProperties = PlatformNumberingProperties.New(xmlElement);

                return numberingProperties;
            }
        }

        private IParagraphStyleId paragraphStyleId;
        public IParagraphStyleId ParagraphStyleId
        {
            get
            {
                if (paragraphStyleId == null)
                    paragraphStyleId = PlatformParagraphStyleId.New(xmlElement);

                return paragraphStyleId;
            }
        }

        private ISpacingBetweenLines spacingBetweenLines;
        public ISpacingBetweenLines SpacingBetweenLines
        {
            get
            {
                if (spacingBetweenLines == null)
                    spacingBetweenLines = PlatformSpacingBetweenLines.New(xmlElement);

                return spacingBetweenLines;
            }
        }

        #region Static helpers methods

        public static PlatformParagraphProperties New(Paragraph paragraph)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<ParagraphProperties>(paragraph);
            return new PlatformParagraphProperties(xmlElement);
        }

        #endregion
    }
}