using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word
{
    public class PlatformRunProperties : PlatformOpenXmlElement, IRunProperties
    {
        private readonly RunProperties xmlElement;

        public PlatformRunProperties()
            : this(new RunProperties())
        {
        }

        public PlatformRunProperties(RunProperties run)
            : base(run)
        {
            this.xmlElement = run;
        }

        #region Interface :

        private IRunFonts runFonts;
        public IRunFonts RunFonts
        {
            get
            {
                if (runFonts == null)
                    runFonts = PlatformRunFonts.New(xmlElement);

                return runFonts;
            }
        }

        public string FontSize
        {
            get
            {
                if (xmlElement.FontSize == null)
                    return null;
                else
                    return xmlElement.FontSize.Val;
            }
            set
            {
                if (value == null)
                    xmlElement.FontSize = null;
                else
                {
                    if (xmlElement.FontSize == null)
                        xmlElement.FontSize = new FontSize();
                    xmlElement.FontSize.Val = value;
                }
            }
        }

        public string Color
        {
            get
            {
                if (xmlElement.Color == null)
                    return null;
                else
                    return xmlElement.Color.Val;
            }
            set
            {
                if (value == null)
                    xmlElement.Color = null;
                else
                {
                    if (xmlElement.Color == null)
                        xmlElement.Color = new Color();
                    xmlElement.Color.Val = value;
                }
            }
        }

        public bool? Bold
        {
            get
            {
                if (xmlElement.Bold == null)
                    return null;
                else
                    return xmlElement.Bold.Val;
            }
            set
            {
                if (value == null)
                    xmlElement.Bold = null;
                else
                {
                    if (xmlElement.Bold == null)
                        xmlElement.Bold = new Bold();
                    xmlElement.Bold.Val = value;
                }
            }
        }

        public bool? Italic
        {
            get
            {
                if (xmlElement.Italic == null)
                    return null;
                else
                    return xmlElement.Italic.Val;
            }
            set
            {
                if (value == null)
                    xmlElement.Italic = null;
                else
                {
                    if (xmlElement.Italic == null)
                        xmlElement.Italic = new Italic();
                    xmlElement.Italic.Val = value;
                }
            }
        }

        public bool? ItalicComplexScript
        {
            get
            {
                if (xmlElement.ItalicComplexScript == null)
                    return null;
                else
                    return xmlElement.ItalicComplexScript.Val;
            }
            set
            {
                if (value == null)
                    xmlElement.ItalicComplexScript = null;
                else
                {
                    if (xmlElement.ItalicComplexScript == null)
                        xmlElement.ItalicComplexScript = new ItalicComplexScript();
                    xmlElement.ItalicComplexScript.Val = value;
                }
            }
        }

        public bool? Caps
        {
            get
            {
                if (xmlElement.Caps == null)
                    return null;
                else
                    return xmlElement.Caps.Val;
            }
            set
            {
                if (value == null)
                    xmlElement.Caps = null;
                else
                {
                    if (xmlElement.Caps == null)
                        xmlElement.Caps = new Caps();
                    xmlElement.Caps.Val = value;
                }
            }
        }

        public bool? DoubleStrike
        {
            get
            {
                if (xmlElement.DoubleStrike == null)
                    return null;
                else
                    return xmlElement.DoubleStrike.Val;
            }
            set
            {
                if (value == null)
                    xmlElement.DoubleStrike = null;
                else
                {
                    if (xmlElement.DoubleStrike == null)
                        xmlElement.DoubleStrike = new DoubleStrike();
                    xmlElement.DoubleStrike.Val = value;
                }
            }
        }

        public bool? Emboss
        {
            get
            {
                if (xmlElement.Emboss == null)
                    return null;
                else
                    return xmlElement.Emboss.Val;
            }
            set
            {
                if (value == null)
                    xmlElement.Emboss = null;
                else
                {
                    if (xmlElement.Emboss == null)
                        xmlElement.Emboss = new Emboss();
                    xmlElement.Emboss.Val = value;
                }
            }
        }

        public bool? NoProof
        {
            get
            {
                if (xmlElement.NoProof == null)
                    return null;
                else
                    return xmlElement.NoProof.Val;
            }
            set
            {
                if (value == null)
                    xmlElement.NoProof = null;
                else
                {
                    if (xmlElement.NoProof == null)
                        xmlElement.NoProof = new NoProof();
                    xmlElement.NoProof.Val = value;
                }
            }
        }

        public bool? Outline
        {
            get
            {
                if (xmlElement.Outline == null)
                    return null;
                else
                    return xmlElement.Outline.Val;
            }
            set
            {
                if (value == null)
                    xmlElement.Outline = null;
                else
                {
                    if (xmlElement.Outline == null)
                        xmlElement.Outline = new Outline();
                    xmlElement.Outline.Val = value;
                }
            }
        }

        public bool? Shadow
        {
            get
            {
                if (xmlElement.Shadow == null)
                    return null;
                else
                    return xmlElement.Shadow.Val;
            }
            set
            {
                if (value == null)
                    xmlElement.Shadow = null;
                else
                {
                    if (xmlElement.Shadow == null)
                        xmlElement.Shadow = new Shadow();
                    xmlElement.Shadow.Val = value;
                }
            }
        }

        #endregion

        #region Static helpers methods

        public static PlatformRunProperties New(Run run)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<RunProperties>(run);
            return new PlatformRunProperties(xmlElement);
        }

        #endregion
    }
}
