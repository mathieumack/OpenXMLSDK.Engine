using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word.Tables;

namespace MvvX.Open_XML_SDK.Shared.Word.Tables
{
    public class PlatformTableCell : PlatformOpenXmlElement, ITableCell
    {
        private readonly TableCell tableCell;

        private ITableCellProperties properties;
        public ITableCellProperties Properties
        {
            get
            {
                return properties;
            }
        }

        public PlatformTableCell(TableCell tableCell)
            : base(tableCell)
        {
            this.tableCell = tableCell;
            this.properties = PlatformTableCellProperties.New();
            tableCell.Append(Properties.ContentItem as TableCellProperties);
        }

        #region Static helpers methods

        public static PlatformTableCell New()
        {
            return new PlatformTableCell(new TableCell());
        }

        #endregion

    }
}
