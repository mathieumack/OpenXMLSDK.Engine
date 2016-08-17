using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.Tables;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Tables
{
    public class PlatformTableStyle : PlatformOpenXmlElement, ITableStyle
    {
        private readonly TableStyle xmlElement;

        public PlatformTableStyle()
            : this(new TableStyle())
        {
        }

        public PlatformTableStyle(TableStyle tableStyle)
            : base(tableStyle)
        {
            this.xmlElement = tableStyle;
        }

        #region Interface :
        
        public string Val
        {
            get
            {
                return xmlElement.Val;
            }

            set
            {
                xmlElement.Val = value;
            }
        }

        #endregion

        #region Static helpers methods
        
        public static PlatformTableStyle New(TableProperties tableProperties)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<TableStyle>(tableProperties);
            return new PlatformTableStyle(xmlElement);
        }

        #endregion
    }
}
