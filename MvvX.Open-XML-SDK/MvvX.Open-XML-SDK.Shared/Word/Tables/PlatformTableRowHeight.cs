using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word.Tables;
using System;

namespace MvvX.Open_XML_SDK.Shared.Word.Tables
{
    public class PlatformTableRowHeight : PlatformOpenXmlElement, ITableRowHeight
    {
        private readonly TableRowHeight xmlElement;

        public PlatformTableRowHeight(TableRowHeight tableWidth)
            : base(tableWidth)
        {
            this.xmlElement = tableWidth;
        }

        #region Interface :
        
        public int? Val
        {
            get
            {
                if (xmlElement.Val.HasValue)
                    return (int?)xmlElement.Val.Value;
                else
                    return null;
            }

            set
            {
                if (value.HasValue)
                    xmlElement.Val = Convert.ToUInt32(value.Value);
                else
                    xmlElement.Val = null;
            }
        }
        
        #endregion

        #region Static helpers methods
        
        public static PlatformTableRowHeight New(TableRowProperties tableProperties)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<TableRowHeight>(tableProperties);
            return new PlatformTableRowHeight(xmlElement);
        }

        #endregion
    }
}
