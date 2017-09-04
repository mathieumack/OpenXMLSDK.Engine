using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word;
using MvvX.Plugins.OpenXMLSDK.Word.Paragraphs;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Paragraphs
{
    public class PlatformParagraphBorders : PlatformOpenXmlElement, IParagraphBorders
    {
        private readonly ParagraphBorders openXmlElement;

        public PlatformParagraphBorders()
            : this(new ParagraphBorders())
        {
        }

        public PlatformParagraphBorders(ParagraphBorders paragraphBorders)
            : base(paragraphBorders)
        {
            this.openXmlElement = paragraphBorders;
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

        public static PlatformParagraphBorders New(ParagraphProperties paragraph)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<ParagraphBorders>(paragraph);
            return new PlatformParagraphBorders(xmlElement);
        }

        #endregion
    }
}
