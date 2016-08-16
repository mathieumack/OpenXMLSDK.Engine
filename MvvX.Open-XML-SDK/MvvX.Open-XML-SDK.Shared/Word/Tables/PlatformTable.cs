using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word.Tables;

namespace MvvX.Open_XML_SDK.Shared.Word.Tables
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

        public PlatformTable(Table table)
            : base(table)
        {
            this.xmlElement = table;
        }

        #region Static helpers methods

        public static PlatformTable New()
        {
            return new PlatformTable(new Table());
        }

        #endregion
    }
}
