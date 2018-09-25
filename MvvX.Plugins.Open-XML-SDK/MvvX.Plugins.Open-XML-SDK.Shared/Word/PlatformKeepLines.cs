using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word
{
    public class PlatformKeepLines : PlatformOpenXmlElement, IKeepLines
    {
        private readonly KeepLines xmlElement;

        public PlatformKeepLines()
            : this(new KeepLines())
        {
        }

        public PlatformKeepLines(KeepLines runFonts)
            : base(runFonts)
        {
            this.xmlElement = runFonts;
        }

        #region Interface


        public OpenXMLSDK.Word.OnOffOnlyValues? Val
        {
            get
            {
                if (xmlElement.Val == null || !xmlElement.Val.HasValue)
                    return null;
                else if(xmlElement.Val.Value)
                    return OpenXMLSDK.Word.OnOffOnlyValues.On;
                else
                    return OpenXMLSDK.Word.OnOffOnlyValues.Off;
            }
            set
            {
                if (value == null || !value.HasValue)
                    xmlElement.Val = null;
                else
                {
                    if (xmlElement.Val == null)
                        xmlElement.Val = new OnOffValue(false);
                    xmlElement.Val.Value = value.Value;
                }
            }
        }


        #endregion

        #region Static helpers methods

        public static PlatformKeepLines New(ParagraphProperties paragraphProperties)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<KeepLines>(paragraphProperties);
            return new PlatformKeepLines(xmlElement);
        }

        #endregion
    }
}
