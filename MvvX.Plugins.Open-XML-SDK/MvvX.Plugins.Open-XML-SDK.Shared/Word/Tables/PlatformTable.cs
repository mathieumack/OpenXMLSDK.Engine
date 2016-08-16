using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.Open_XML_SDK.Core.Word.Tables;

namespace MvvX.Plugins.Open_XML_SDK.Shared.Word.Tables
{
    public class PlatformTable : PlatformOpenXmlElement, ITable
    {
        private readonly Table xmlElement;

        private ITableProperties properties;
        public ITableProperties Properties
        {
            get
            {
                if(properties == null)
                    properties = PlatformTableProperties.New(xmlElement);
                return properties;
            }
        }

        public PlatformTable()
            : this(new Table())
        {
        }

        public PlatformTable(Table table)
            : base(table)
        {
            this.xmlElement = table;
        }
    }
}
