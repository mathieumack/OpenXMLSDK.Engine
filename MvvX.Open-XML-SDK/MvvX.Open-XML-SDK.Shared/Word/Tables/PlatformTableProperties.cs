using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word.Bases;
using MvvX.Open_XML_SDK.Core.Word.Tables.Models;

namespace MvvX.Open_XML_SDK.Shared.Word.Tables
{
    public class PlatformTableProperties : PlatformOpenXmlElement, ITableProperties
    {
        private readonly TableProperties tableProperties;
       

        public PlatformTableProperties(TableProperties tableProperties)
            : base(tableProperties)
        {
            this.tableProperties = tableProperties;
        }

        public PlatformTableProperties(params OpenXmlElement[] childElements)
        {
            this.ChildItem = childElements;
        }
    }
}
