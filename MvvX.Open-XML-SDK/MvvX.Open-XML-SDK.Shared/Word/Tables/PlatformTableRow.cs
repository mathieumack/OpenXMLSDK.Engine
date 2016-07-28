using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word.Bases;

namespace MvvX.Open_XML_SDK.Core.Word.Tables
{
    public class PlatformTableRow : PlatformOpenXmlElement, ITableRow
    {
        private readonly TableRow tableRow;

        public PlatformTableRow(TableRow tableRow)
            : base(table)
        {
            this.tableRow = tableRow;
        }
    }
}
