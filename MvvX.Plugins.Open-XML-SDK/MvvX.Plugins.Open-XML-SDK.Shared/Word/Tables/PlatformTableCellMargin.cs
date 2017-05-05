using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.Tables;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Tables
{
    public class PlatformTableCellMargin : PlatformOpenXmlElement, ITableCellMargin
    {
        private readonly TableCellMargin openXmlElement;

        public PlatformTableCellMargin(OpenXmlElement openXmlElement) : base(openXmlElement)
        {
        }

        public PlatformTableCellMargin(TableCellMargin tableCellMargin)
            : base(tableCellMargin)
        {
            this.openXmlElement = tableCellMargin;
        }

        private ITableWidth bottomMargin;
        public ITableWidth BottomMargin
        {
            get
            {
                if (bottomMargin == null)
                    bottomMargin = PlatformTableCellMarginWidth<BottomMargin>.New(openXmlElement);

                return bottomMargin;
            }
        }

        private ITableWidth leftMargin;
        public ITableWidth LeftMargin
        {
            get
            {
                if (leftMargin == null)
                    leftMargin = PlatformTableCellMarginWidth<LeftMargin>.New(openXmlElement);

                return leftMargin;
            }
        }

        private ITableWidth rightMargin;
        public ITableWidth RightMargin
        {
            get
            {
                if (rightMargin == null)
                    rightMargin = PlatformTableCellMarginWidth<RightMargin>.New(openXmlElement);

                return rightMargin;
            }
        }

        private ITableWidth topMargin;
        public ITableWidth TopMargin
        {
            get
            {
                if (topMargin == null)
                    topMargin = PlatformTableCellMarginWidth<TopMargin>.New(openXmlElement);

                return topMargin;
            }
        }

        #region Static helpers methods

        public static PlatformTableCellMargin New(TableCellProperties tableProperties)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<TableCellMargin>(tableProperties);
            return new PlatformTableCellMargin(xmlElement);
        }

        #endregion
    }
}
