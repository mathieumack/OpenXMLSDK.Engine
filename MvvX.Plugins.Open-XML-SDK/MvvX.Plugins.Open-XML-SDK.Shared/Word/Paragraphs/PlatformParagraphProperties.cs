using System;
using System.Collections.Generic;
using System.Text;
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

        #region Static helpers methods

        public static PlatformParagraphProperties New(Paragraph paragraph)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<ParagraphProperties>(paragraph);
            return new PlatformParagraphProperties(xmlElement);
        }

        #endregion
    }
}
