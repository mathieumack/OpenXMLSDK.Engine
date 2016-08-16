using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word.Tables;
using MvvX.Open_XML_SDK.Core.Word;
using System.Linq;

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

        private IBorderType insideVerticalBorder;
        public IBorderType InsideVerticalBorder
        {
            get
            {
                if (insideVerticalBorder == null)
                    insideVerticalBorder = PlatformBorder<InsideVerticalBorder>.New(openXmlElement);

                return insideVerticalBorder;
            }
        }

        private IBorderType insideHorizontalBorder;
        public IBorderType InsideHorizontalBorder
        {
            get
            {
                if (insideHorizontalBorder == null)
                    insideHorizontalBorder = PlatformBorder<InsideHorizontalBorder>.New(openXmlElement);

                return insideHorizontalBorder;
            }
        }
        
        #endregion

        #region Static helpers methods
        
        public static PlatformTableBorders New(TableProperties tableProperties)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<TableBorders>(tableProperties);
            return new PlatformTableBorders(xmlElement);
        }

        #endregion
    }
}
