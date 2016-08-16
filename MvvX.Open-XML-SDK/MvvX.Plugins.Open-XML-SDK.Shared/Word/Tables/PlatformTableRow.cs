using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.Open_XML_SDK.Core.Word.Tables;

namespace MvvX.Plugins.Open_XML_SDK.Shared.Word.Tables
{
    public class PlatformTableRow : PlatformOpenXmlElement, ITableRow
    {
        private readonly TableRow xmlElement;

        private ITableRowProperties properties;
        public ITableRowProperties Properties
        {
            get
            {
                if (properties == null)
                    properties = PlatformTableRowProperties.New(xmlElement);
                return properties;
            }
        }

        public PlatformTableRow()
            : this(new TableRow())
        {
        }

        public PlatformTableRow(TableRow tableRow)
            : base(tableRow)
        {
            this.xmlElement = tableRow;
        }

        #region Static helpers methods

        public static PlatformTableRow New()
        {
            return new PlatformTableRow(new TableRow());
        }

        #endregion
    }
}
