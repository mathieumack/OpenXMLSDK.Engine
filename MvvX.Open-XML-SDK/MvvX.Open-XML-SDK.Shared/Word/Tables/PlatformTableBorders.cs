using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word.Tables;
using MvvX.Open_XML_SDK.Core.Word;

namespace MvvX.Open_XML_SDK.Shared.Word.Tables
{
    public class PlatformTableBorders : PlatformOpenXmlElement, ITableBorders
    {
        private readonly TableBorders openXmlElement;

        public PlatformTableBorders(TableBorders tableBorders)
            : base(tableBorders)
        {
            this.openXmlElement = tableBorders;
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

        private IBorderType insideVerticalBorder;
        public IBorderType InsideVerticalBorder
        {
            get
            {
                if (insideVerticalBorder == null)
                    insideVerticalBorder = PlatformBorder<InsideVerticalBorder>.New(openXmlElement.InsideVerticalBorder);

                return insideVerticalBorder;
            }
        }

        private IBorderType insideHorizontalBorder;
        public IBorderType InsideHorizontalBorder
        {
            get
            {
                if (insideHorizontalBorder == null)
                    insideHorizontalBorder = PlatformBorder<InsideHorizontalBorder>.New(openXmlElement.InsideHorizontalBorder);

                return insideHorizontalBorder;
            }
        }
        
        #endregion

        #region Static helpers methods

        public static PlatformTableBorders New()
        {
            return new PlatformTableBorders(new TableBorders());
        }

        public static PlatformTableBorders New(TableBorders tableBorders)
        {
            if (tableBorders == null)
                tableBorders = new TableBorders();

            return new PlatformTableBorders(tableBorders);
        }

        #endregion
    }
}
