using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.Tables;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Tables
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
