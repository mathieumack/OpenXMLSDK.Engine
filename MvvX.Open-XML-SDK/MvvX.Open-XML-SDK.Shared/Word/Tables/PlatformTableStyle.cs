using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word.Bases;
using MvvX.Open_XML_SDK.Core.Word.Tables;

namespace MvvX.Open_XML_SDK.Shared.Word.Tables
{
    public class PlatformTableStyle : PlatformOpenXmlElement, ITableStyle
    {
        private readonly TableStyle tableStyle;

        public PlatformTableStyle(TableStyle tableStyle)
            : base(tableStyle)
        {
            this.tableStyle = tableStyle;
        }
    }
}
