using System;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word
{
    public class PlatformRunFonts : PlatformOpenXmlElement, IRunFonts
    {
        private readonly RunFonts xmlElement;

        public PlatformRunFonts()
            : this(new RunFonts())
        {
        }

        public PlatformRunFonts(RunFonts runFonts)
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

        public string ComplexScript
        {
            get
            {
                if (xmlElement.ComplexScript == null)
                    return null;
                else
                    return xmlElement.ComplexScript.Value;
            }
            set
            {
                if (value == null)
                    xmlElement.ComplexScript = null;
                else
                {
                    if (xmlElement.ComplexScript == null)
                        xmlElement.ComplexScript = new DocumentFormat.OpenXml.StringValue();
                    xmlElement.ComplexScript.Value = value;
                }
            }
        }

        public string EastAsia
        {
            get
            {
                if (xmlElement.EastAsia == null)
                    return null;
                else
                    return xmlElement.EastAsia.Value;
            }
            set
            {
                if (value == null)
                    xmlElement.EastAsia = null;
                else
                {
                    if (xmlElement.EastAsia == null)
                        xmlElement.EastAsia = new DocumentFormat.OpenXml.StringValue();
                    xmlElement.EastAsia.Value = value;
                }
            }
        }

        public string HighAnsi
        {
            get
            {
                if (xmlElement.HighAnsi == null)
                    return null;
                else
                    return xmlElement.HighAnsi.Value;
            }
            set
            {
                if (value == null)
                    xmlElement.HighAnsi = null;
                else
                {
                    if (xmlElement.HighAnsi == null)
                        xmlElement.HighAnsi = new DocumentFormat.OpenXml.StringValue();
                    xmlElement.HighAnsi.Value = value;
                }
            }
        }

        #endregion

        #region Static helpers methods
        
        public static PlatformRunFonts New(RunProperties runProperties)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<RunFonts>(runProperties);
            return new PlatformRunFonts(xmlElement);
        }

        #endregion
    }
}
