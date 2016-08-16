using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word.Tables;

namespace MvvX.Open_XML_SDK.Shared.Word.Tables
{
    public class PlatformTableStyle : PlatformOpenXmlElement, ITableStyle
    {
        private readonly TableStyle xmlElement;

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
