using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.Open_XML_SDK.Core.Word.Tables;

namespace MvvX.Plugins.Open_XML_SDK.Shared.Word.Tables
{
    public class PlatformTableCellWidth : PlatformOpenXmlElement, ITableCellWidth
    {
        private readonly TableCellWidth xmlElement;

        public PlatformTableCellWidth()
            : this(new TableCellWidth())
        {
        }

        public PlatformTableCellWidth(TableCellWidth tableWidth)
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
        
        public static PlatformTableCellWidth New(TableCellProperties tableProperties)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<TableCellWidth>(tableProperties);
            return new PlatformTableCellWidth(xmlElement);
        }

        #endregion
    }
}
