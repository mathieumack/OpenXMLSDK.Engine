using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word;
using MvvX.Open_XML_SDK.Core.Word.Tables;

namespace MvvX.Open_XML_SDK.Shared.Word.Tables
{
    public class PlatformTableCellProperties : PlatformOpenXmlElement, ITableCellProperties
    {
        private readonly TableCellProperties xmlElement;

        public PlatformTableCellProperties(TableCellProperties tableCell)
            : base(tableCell)
        {
            this.xmlElement = tableCell;
        }

        #region Interface :

        private ITableCellBorders tableCellProperties;
        public ITableCellBorders TableCellProperties
        {
            get
            {
                if (tableCellProperties == null)
                    tableCellProperties = PlatformTableCellBorders.New(xmlElement.TableCellBorders);

                return tableCellProperties;
            }
        }

        private IGridSpan gridSpan;
        public IGridSpan GridSpan
        {
            get
            {
                if (gridSpan == null)
                    gridSpan = PlatformGridSpan.New(xmlElement.GridSpan);
                return gridSpan;
            }
        }

        public Core.Word.OnOffOnlyValues? NoWrap
        {
            get
            {
                if (xmlElement.NoWrap == null || !xmlElement.NoWrap.Val.HasValue)
                    return null;
                else
                    return (Core.Word.OnOffOnlyValues)(int)xmlElement.NoWrap.Val.Value;
            }
            set
            {
                if (value == null)
                    xmlElement.NoWrap = null;
                else
                {
                    if (xmlElement.NoWrap == null)
                        xmlElement.NoWrap = new NoWrap();
                    xmlElement.NoWrap.Val = (DocumentFormat.OpenXml.Wordprocessing.OnOffOnlyValues)(int)value;
                }
            }
        }

        #endregion

        #region Static helpers methods

        public static PlatformTableCellProperties New()
        {
            return new PlatformTableCellProperties(new TableCellProperties());
        }

        #endregion
    }
}
