using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word.Bases;

namespace MvvX.Open_XML_SDK.Core.Word.Tables
{
    public class PlatformTableCell : PlatformOpenXmlElement, ITableCell
    {
        private readonly TableCell tableCell;

        public PlatformTableCell(TableCell tableCell)
            : base(table)
        {
            this.tableCell = tableCell;
        }
    }
}
