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
                    topBorder = PlatformBorder<TopBorder>.New(openXmlElement);

                return topBorder;
            }
        }

        private IBorderType rightBorder;
        public IBorderType RightBorder
        {
            get
            {
                if (rightBorder == null)
                    rightBorder = PlatformBorder<RightBorder>.New(openXmlElement);

                return rightBorder;
            }
        }

        private IBorderType leftBorder;
        public IBorderType LeftBorder
        {
            get
            {
                if (leftBorder == null)
                    leftBorder = PlatformBorder<LeftBorder>.New(openXmlElement);

                return leftBorder;
            }
        }

        private IBorderType bottomBorder;
        public IBorderType BottomBorder
        {
            get
            {
                if (bottomBorder == null)
                    bottomBorder = PlatformBorder<BottomBorder>.New(openXmlElement);

                return bottomBorder;
            }
        }

        private IBorderType topLeftToBottomRightCellBorder;
        public IBorderType TopLeftToBottomRightCellBorder
        {
            get
            {
                if (topLeftToBottomRightCellBorder == null)
                    topLeftToBottomRightCellBorder = PlatformBorder<TopLeftToBottomRightCellBorder>.New(openXmlElement);

                return topLeftToBottomRightCellBorder;
            }
        }

        private IBorderType topRightToBottomLeftCellBorder;
        public IBorderType TopRightToBottomLeftCellBorder
        {
            get
            {
                if (topRightToBottomLeftCellBorder == null)
                    topRightToBottomLeftCellBorder = PlatformBorder<TopRightToBottomLeftCellBorder>.New(openXmlElement);

                return topRightToBottomLeftCellBorder;
            }
        }

        #endregion

        #region Static helpers methods

        public static PlatformTableCellBorders New(TableCellProperties parent)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<TableCellBorders>(parent);
            return new PlatformTableCellBorders(xmlElement);
        }

        #endregion
    }
}
