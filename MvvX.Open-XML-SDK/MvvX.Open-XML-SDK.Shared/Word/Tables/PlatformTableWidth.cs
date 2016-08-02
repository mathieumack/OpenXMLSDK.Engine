using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word.Bases;
using MvvX.Open_XML_SDK.Core.Word.Tables;

namespace MvvX.Open_XML_SDK.Shared.Word.Tables
{
    public class PlatformTableWidth : PlatformOpenXmlElement, ITableWidth
    {
        private readonly TableWidth tableWidth;

        public PlatformTableWidth(TableWidth tableWidth)
            : base(tableWidth)
        {
            this.tableWidth = tableWidth;
        }
    }
}
