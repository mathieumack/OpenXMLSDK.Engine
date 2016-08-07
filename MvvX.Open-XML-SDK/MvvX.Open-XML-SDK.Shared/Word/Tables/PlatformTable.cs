using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word.Tables;

namespace MvvX.Open_XML_SDK.Shared.Word.Tables
{
    public class PlatformTable : PlatformOpenXmlElement, ITable
    {
        private readonly Table table;

        private ITableProperties properties;
        public ITableProperties Properties
        {
            get
            {
                return properties;
            }
        }

        public PlatformTable(Table table)
            : base(table)
        {
            this.table = table;
            this.properties = PlatformTableProperties.New();
            table.Append(Properties.ContentItem as TableProperties);
        }

        #region Static helpers methods

        public static PlatformTable New()
        {
            return new PlatformTable(new Table());
        }

        #endregion
    }
}
