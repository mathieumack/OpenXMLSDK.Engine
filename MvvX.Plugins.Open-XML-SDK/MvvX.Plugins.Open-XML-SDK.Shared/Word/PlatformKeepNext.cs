using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word
{
    public class PlatformKeepNext : PlatformOpenXmlElement, IKeepNext
    {
        private readonly KeepNext xmlElement;

        public PlatformKeepNext()
            : this(new KeepNext())
        {
        }

        public PlatformKeepNext(KeepNext runFonts)
            : base(runFonts)
        {
            this.xmlElement = runFonts;
        }

        #region Interface

        public string Ascii
        {
            get
            {
                if (xmlElement.Ascii == null)
                    return null;
                else
                    return xmlElement.Ascii.Value;
            }
            set
            {
                if (value == null)
                    xmlElement.Ascii = null;
                else
                {
                    if (xmlElement.Ascii == null)
                        xmlElement.Ascii = new DocumentFormat.OpenXml.StringValue();
                    xmlElement.Ascii.Value = value;
                }
            }
        }

        #endregion

        #region Static helpers methods

        public static PlatformKeepNext New(ParagraphProperties paragraphProperties)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<KeepNext>(paragraphProperties);
            return new PlatformKeepNext(xmlElement);
        }

        #endregion
    }
