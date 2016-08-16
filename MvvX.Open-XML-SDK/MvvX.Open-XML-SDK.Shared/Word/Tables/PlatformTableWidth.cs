using System;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word;
using MvvX.Open_XML_SDK.Core.Word.Tables;
using System.Linq;

namespace MvvX.Open_XML_SDK.Shared.Word.Tables
{
    public class PlatformTableWidth : PlatformOpenXmlElement, ITableWidth
    {
        private readonly TableWidth xmlElement;

        public PlatformTableWidth(TableWidth tableWidth)
            : base(tableWidth)
        {
            this.xmlElement = tableWidth;
        }

        #region Interface :
        
        public string Width
        {
            get
            {
                return xmlElement.Width;
            }

            set
            {
                xmlElement.Width = value;
            }
        }

        public Core.Word.Tables.TableWidthUnitValues? Type
        {
            get
            {
                if (xmlElement.Type.HasValue)
                    return (Core.Word.Tables.TableWidthUnitValues)(int)xmlElement.Type.Value;
                else
                    return null;
            }

            set
            {
                if (!value.HasValue)
                    xmlElement.Type = null;
                else
                    xmlElement.Type = (DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues)(int)value.Value;
            }
        }

        #endregion

        #region Static helpers methods
        
        public static PlatformTableWidth New(TableProperties tableProperties)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<TableWidth>(tableProperties);
            return new PlatformTableWidth(xmlElement);
        }

        #endregion
    }
}
