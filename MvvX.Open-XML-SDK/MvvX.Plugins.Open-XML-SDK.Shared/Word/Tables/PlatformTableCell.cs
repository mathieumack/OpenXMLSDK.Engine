using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.Open_XML_SDK.Core.Word.Tables;

namespace MvvX.Plugins.Open_XML_SDK.Shared.Word.Tables
{
    public class PlatformTableCell : PlatformOpenXmlElement, ITableCell
    {
        private readonly TableCell xmlElement;

        private ITableCellProperties properties;
        public ITableCellProperties Properties
        {
            get
            {
                if (properties == null)
                    properties = PlatformTableCellProperties.New(xmlElement);
                return properties;
            }
        }

        public PlatformTableCell()
            : this(new TableCell())
        {
        }

        public PlatformTableCell(TableCell tableCell)
            : base(tableCell)
        {
            this.xmlElement = tableCell;
        }

        #region Static helpers methods

        public static PlatformTableCell New()
        {
            return new PlatformTableCell(new TableCell());
        }

        #endregion

    }
}
