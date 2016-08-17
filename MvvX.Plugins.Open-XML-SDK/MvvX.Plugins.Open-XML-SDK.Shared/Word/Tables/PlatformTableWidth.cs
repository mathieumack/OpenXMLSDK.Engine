using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.Tables;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Tables
{
    public class PlatformTableWidth : PlatformOpenXmlElement, ITableWidth
    {
        private readonly TableWidth xmlElement;

        public PlatformTableWidth()
            : this(new TableWidth())
        {
        }

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

        public OpenXMLSDK.Word.Tables.TableWidthUnitValues? Type
        {
            get
            {
                if (xmlElement.Type.HasValue)
                    return (OpenXMLSDK.Word.Tables.TableWidthUnitValues)(int)xmlElement.Type.Value;
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
