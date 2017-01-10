using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.Paragraphs;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Paragraphs
{
    public class PlatformNumberingProperties : PlatformOpenXmlElement, INumberingProperties
    {
        private readonly NumberingProperties xmlElement;

        public PlatformNumberingProperties()
            : this(new NumberingProperties())
        {
        }

        public PlatformNumberingProperties(NumberingProperties numberingProp)
            : base(numberingProp)
        {
            this.xmlElement = numberingProp;
        }

        public int? NumberingLevelReference
        {
            get
            {
                if (xmlElement.NumberingLevelReference == null)
                    return null;
                else
                    return xmlElement.NumberingLevelReference.Val;
            }
            set
            {
                if (value == null)
                    xmlElement.NumberingLevelReference = null;
                else
                {
                    if (xmlElement.NumberingLevelReference == null)
                        xmlElement.NumberingLevelReference = new NumberingLevelReference();
                    xmlElement.NumberingLevelReference.Val = value;
                }
            }
        }

        public int? NumberingId
        {
            get
            {
                if (xmlElement.NumberingId == null)
                    return null;
                else
                    return xmlElement.NumberingId.Val;
            }
            set
            {
                if (value == null)
                    xmlElement.NumberingId = null;
                else
                {
                    if (xmlElement.NumberingId == null)
                        xmlElement.NumberingId = new NumberingId();
                    xmlElement.NumberingId.Val = value;
                }
            }
        }

        #region Static helpers methods

        public static PlatformNumberingProperties New(ParagraphProperties paragraphProp)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<NumberingProperties>(paragraphProp);
            return new PlatformNumberingProperties(xmlElement);
        }

        #endregion
    }
}
