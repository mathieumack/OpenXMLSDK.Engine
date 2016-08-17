using System;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word
{
    public class PlatformShading : PlatformOpenXmlElement, IShading
    {
        private readonly Shading xmlElement;

        public PlatformShading()
            : this(new Shading())
        {
        }

        public PlatformShading(Shading shading)
            : base(shading)
        {
            this.xmlElement = shading;
        }

        #region Interface

        public string Color
        {
            get
            {
                return xmlElement.Color;
            }

            set
            {
                xmlElement.Color = value;
            }
        }

        public string Fill
        {
            get
            {
                return xmlElement.Fill;
            }

            set
            {
                xmlElement.Fill = value;
            }
        }

        public OpenXMLSDK.Word.ThemeColorValues? ThemeColor
        {
            get
            {
                if (xmlElement.ThemeColor != null && xmlElement.ThemeColor.HasValue)
                    return (OpenXMLSDK.Word.ThemeColorValues)(int)xmlElement.ThemeColor.Value;
                else
                    return null;
            }

            set
            {
                if (value.HasValue)
                    xmlElement.ThemeColor = (DocumentFormat.OpenXml.Wordprocessing.ThemeColorValues)(int)value;
                else
                    xmlElement.ThemeColor = null;
            }
        }

        public OpenXMLSDK.Word.ThemeColorValues? ThemeFill
        {
            get
            {
                if (xmlElement.ThemeFill != null && xmlElement.ThemeFill.HasValue)
                    return (OpenXMLSDK.Word.ThemeColorValues)(int)xmlElement.ThemeFill.Value;
                else
                    return null;
            }

            set
            {
                if (value.HasValue)
                    xmlElement.ThemeFill = (DocumentFormat.OpenXml.Wordprocessing.ThemeColorValues)(int)value;
                else
                    xmlElement.ThemeFill = null;
            }
        }

        public string ThemeFillShade
        {
            get
            {
                return xmlElement.ThemeFillShade;
            }

            set
            {
                xmlElement.ThemeFillShade = value;
            }
        }

        public string ThemeFillTint
        {
            get
            {
                return xmlElement.ThemeFillTint;
            }

            set
            {
                xmlElement.ThemeFillTint = value;
            }
        }

        public string ThemeShade
        {
            get
            {
                return xmlElement.ThemeShade;
            }

            set
            {
                xmlElement.ThemeShade = value;
            }
        }

        public string ThemeTint
        {
            get
            {
                return xmlElement.ThemeTint;
            }

            set
            {
                xmlElement.ThemeTint = value;
            }
        }

        public OpenXMLSDK.Word.ShadingPatternValues? Val
        {
            get
            {
                if (xmlElement.Val != null && xmlElement.Val.HasValue)
                    return (OpenXMLSDK.Word.ShadingPatternValues)(int)xmlElement.Val.Value;
                else
                    return null;
            }

            set
            {
                if (value.HasValue)
                    xmlElement.Val = (DocumentFormat.OpenXml.Wordprocessing.ShadingPatternValues)(int)value;
                else
                    xmlElement.Val = null;
            }
        }

        #endregion

        #region Static helpers methods

        public static PlatformShading New(TableProperties tableProperties)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<Shading>(tableProperties);
            return new PlatformShading(xmlElement);
        }

        public static PlatformShading New(TableCellProperties tableProperties)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<Shading>(tableProperties);
            return new PlatformShading(xmlElement);
        }

        #endregion
    }
}
