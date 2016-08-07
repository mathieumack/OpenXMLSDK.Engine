using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word.Tables;
using System;
using MvvX.Open_XML_SDK.Core.Word;

namespace MvvX.Open_XML_SDK.Shared.Word.Tables
{
    public class PlatformTableCellBorders : PlatformOpenXmlElement, ITableCellBorders
    {
        private readonly TableCellBorders openXmlElement;

        public PlatformTableCellBorders(TableCellBorders tableCellBorders)
            : base(tableCellBorders)
        {
            this.openXmlElement = tableCellBorders;
        }

        #region Interface

        private IBorderType topBorder;
        public IBorderType TopBorder
        {
            get
            {
                if (topBorder == null)
                    topBorder = PlatformBorder<TopBorder>.New(openXmlElement.TopBorder);

                return topBorder;
            }
        }

        private IBorderType rightBorder;
        public IBorderType RightBorder
        {
            get
            {
                if (rightBorder == null)
                    rightBorder = PlatformBorder<RightBorder>.New(openXmlElement.RightBorder);

                return rightBorder;
            }
        }

        private IBorderType leftBorder;
        public IBorderType LeftBorder
        {
            get
            {
                if (leftBorder == null)
                    leftBorder = PlatformBorder<LeftBorder>.New(openXmlElement.LeftBorder);

                return leftBorder;
            }
        }

        private IBorderType bottomBorder;
        public IBorderType BottomBorder
        {
            get
            {
                if (bottomBorder == null)
                    bottomBorder = PlatformBorder<BottomBorder>.New(openXmlElement.BottomBorder);

                return bottomBorder;
            }
        }

        #endregion

        #region Static helpers methods

        public static PlatformTableCellBorders New()
        {
            return new PlatformTableCellBorders(new TableCellBorders());
        }

        public static PlatformTableCellBorders New(TableCellBorders tableCellBorders)
        {
            if (tableCellBorders == null)
                tableCellBorders = new TableCellBorders();

            return new PlatformTableCellBorders(tableCellBorders);
        }

        #endregion
    }
}
