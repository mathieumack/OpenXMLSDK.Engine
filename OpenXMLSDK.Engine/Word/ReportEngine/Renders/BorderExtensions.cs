using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.Word.ReportEngine.Models.ExtendedModels;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    public static class BorderExtensions
    {
        /// <summary>
        /// Render a border for a table
        /// </summary>
        /// <param name="border"></param>
        /// <returns></returns>
        public static TableBorders Render(this BorderModel border)
        {
            TableBorders borders = new TableBorders();

            FillTableBorders(border, borders);

            return borders;
        }

        /// <summary>
        /// Fill TableBorders or TableCellBorders element with borders.
        /// </summary>
        /// <param name="border"></param>
        /// <param name="borders"></param>
        private static void FillTableBorders(BorderModel border, TableBorders borders)
        {
            if (border.BorderPositions.HasFlag(BorderPositions.LEFT))
            {
                var leftBorder = new LeftBorder();
                leftBorder.Color = border.UseVariableBorders && !string.IsNullOrWhiteSpace(border.BorderLeftColor) ? border.BorderLeftColor : border.BorderColor;
                leftBorder.Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Thick;
                leftBorder.Size = border.UseVariableBorders ? border.BorderWidthLeft : border.BorderWidth;
                borders.LeftBorder = leftBorder;
            }

            if (border.BorderPositions.HasFlag(BorderPositions.TOP))
            {
                var topBorder = new TopBorder();
                topBorder.Color = border.UseVariableBorders && !string.IsNullOrWhiteSpace(border.BorderTopColor) ? border.BorderTopColor : border.BorderColor;
                topBorder.Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Thick;
                topBorder.Size = border.UseVariableBorders ? border.BorderWidthTop : border.BorderWidth;
                borders.TopBorder = topBorder;
            }

            if (border.BorderPositions.HasFlag(BorderPositions.RIGHT))
            {
                var rightBorder = new RightBorder();
                rightBorder.Color = border.UseVariableBorders && !string.IsNullOrWhiteSpace(border.BorderRightColor) ? border.BorderRightColor : border.BorderColor;
                rightBorder.Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Thick;
                rightBorder.Size = border.UseVariableBorders ? border.BorderWidthRight : border.BorderWidth;
                borders.RightBorder = rightBorder;
            }

            if (border.BorderPositions.HasFlag(BorderPositions.BOTTOM))
            {
                var bottomBorder = new BottomBorder();
                bottomBorder.Color = border.UseVariableBorders && !string.IsNullOrWhiteSpace(border.BorderBottomColor) ? border.BorderBottomColor : border.BorderColor;
                bottomBorder.Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Thick;
                bottomBorder.Size = border.UseVariableBorders ? border.BorderWidthBottom : border.BorderWidth;
                borders.BottomBorder = bottomBorder;
            }

            if (border.BorderPositions.HasFlag(BorderPositions.INSIDEHORIZONTAL))
            {
                var insideHorizontalBorder = new InsideHorizontalBorder();
                insideHorizontalBorder.Color = border.BorderColor;
                insideHorizontalBorder.Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Thick;
                insideHorizontalBorder.Size = border.UseVariableBorders ? border.BorderWidthInsideHorizontal : border.BorderWidth;
                borders.InsideHorizontalBorder = insideHorizontalBorder;
            }

            if (border.BorderPositions.HasFlag(BorderPositions.INSIDEVERTICAL))
            {
                var insideVerticalBorder = new InsideVerticalBorder();
                insideVerticalBorder.Color = border.BorderColor;
                insideVerticalBorder.Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Thick;
                insideVerticalBorder.Size = border.UseVariableBorders ? border.BorderWidthInsideVertical : border.BorderWidth;
                borders.InsideVerticalBorder = insideVerticalBorder;
            }
        }

        /// <summary>
        /// Render a border for a table cell
        /// </summary>
        /// <param name="border"></param>
        /// <returns></returns>
        public static TableCellBorders RenderCellBorder(this BorderModel border)
        {
            TableCellBorders borders = new TableCellBorders();

            FillTableCellBorders(border, borders);

            return borders;
        }

        /// <summary>
        /// Fill TableBorders or TableCellBorders element with borders.
        /// </summary>
        /// <param name="border"></param>
        /// <param name="borders"></param>
        private static void FillTableCellBorders(BorderModel border, TableCellBorders borders)
        {
            if (border.BorderPositions.HasFlag(BorderPositions.LEFT))
            {
                var leftBorder = new LeftBorder();
                leftBorder.Color = border.UseVariableBorders && !string.IsNullOrWhiteSpace(border.BorderLeftColor) ? border.BorderLeftColor : border.BorderColor;
                leftBorder.Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Thick;
                leftBorder.Size = border.UseVariableBorders ? border.BorderWidthLeft : border.BorderWidth;
                borders.LeftBorder = leftBorder;
            }

            if (border.BorderPositions.HasFlag(BorderPositions.TOP))
            {
                var topBorder = new TopBorder();
                topBorder.Color = border.UseVariableBorders && !string.IsNullOrWhiteSpace(border.BorderTopColor) ? border.BorderTopColor : border.BorderColor;
                topBorder.Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Thick;
                topBorder.Size = border.UseVariableBorders ? border.BorderWidthTop : border.BorderWidth;
                borders.TopBorder = topBorder;
            }

            if (border.BorderPositions.HasFlag(BorderPositions.RIGHT))
            {
                var rightBorder = new RightBorder();
                rightBorder.Color = border.UseVariableBorders && !string.IsNullOrWhiteSpace(border.BorderRightColor) ? border.BorderRightColor : border.BorderColor;
                rightBorder.Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Thick;
                rightBorder.Size = border.UseVariableBorders ? border.BorderWidthRight : border.BorderWidth;
                borders.RightBorder = rightBorder;
            }

            if (border.BorderPositions.HasFlag(BorderPositions.BOTTOM))
            {
                var bottomBorder = new BottomBorder();
                bottomBorder.Color = border.UseVariableBorders && !string.IsNullOrWhiteSpace(border.BorderBottomColor) ? border.BorderBottomColor : border.BorderColor;
                bottomBorder.Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Thick;
                bottomBorder.Size = border.UseVariableBorders ? border.BorderWidthBottom : border.BorderWidth;
                borders.BottomBorder = bottomBorder;
            }

            if (border.BorderPositions.HasFlag(BorderPositions.INSIDEHORIZONTAL))
            {
                var insideHorizontalBorder = new InsideHorizontalBorder();
                insideHorizontalBorder.Color = border.BorderColor;
                insideHorizontalBorder.Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Thick;
                insideHorizontalBorder.Size = border.UseVariableBorders ? border.BorderWidthInsideHorizontal : border.BorderWidth;
                borders.InsideHorizontalBorder = insideHorizontalBorder;
            }

            if (border.BorderPositions.HasFlag(BorderPositions.INSIDEVERTICAL))
            {
                var insideVerticalBorder = new InsideVerticalBorder();
                insideVerticalBorder.Color = border.BorderColor;
                insideVerticalBorder.Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Thick;
                insideVerticalBorder.Size = border.UseVariableBorders ? border.BorderWidthInsideVertical : border.BorderWidth;
                borders.InsideVerticalBorder = insideVerticalBorder;
            }
        }

        /// <summary>
        /// Render borders for a paragraph
        /// </summary>
        /// <param name="border"></param>
        /// <returns></returns>
        public static ParagraphBorders RenderParagraphBorder(this BorderModel border)
        {
            ParagraphBorders borders = new ParagraphBorders();

            FillParagrahBorders(border, borders);

            return borders;
        }

        /// <summary>
        /// Fill TableBorders or TableCellBorders element with borders.
        /// </summary>
        /// <param name="border"></param>
        /// <param name="borders"></param>
        private static void FillParagrahBorders(BorderModel border, ParagraphBorders borders)
        {
            if (border.BorderPositions.HasFlag(BorderPositions.LEFT))
            {
                var leftBorder = new LeftBorder();
                leftBorder.Color = border.UseVariableBorders && !string.IsNullOrWhiteSpace(border.BorderLeftColor) ? border.BorderLeftColor : border.BorderColor;
                leftBorder.Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Thick;
                leftBorder.Size = border.UseVariableBorders ? border.BorderWidthLeft : border.BorderWidth;
                borders.LeftBorder = leftBorder;
            }

            if (border.BorderPositions.HasFlag(BorderPositions.TOP))
            {
                var topBorder = new TopBorder();
                topBorder.Color = border.UseVariableBorders && !string.IsNullOrWhiteSpace(border.BorderTopColor) ? border.BorderTopColor : border.BorderColor;
                topBorder.Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Thick;
                topBorder.Size = border.UseVariableBorders ? border.BorderWidthTop : border.BorderWidth;
                borders.TopBorder = topBorder;
            }

            if (border.BorderPositions.HasFlag(BorderPositions.RIGHT))
            {
                var rightBorder = new RightBorder();
                rightBorder.Color = border.UseVariableBorders && !string.IsNullOrWhiteSpace(border.BorderRightColor) ? border.BorderRightColor : border.BorderColor;
                rightBorder.Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Thick;
                rightBorder.Size = border.UseVariableBorders ? border.BorderWidthRight : border.BorderWidth;
                borders.RightBorder = rightBorder;
            }

            if (border.BorderPositions.HasFlag(BorderPositions.BOTTOM))
            {
                var bottomBorder = new BottomBorder();
                bottomBorder.Color = border.UseVariableBorders && !string.IsNullOrWhiteSpace(border.BorderBottomColor) ? border.BorderBottomColor : border.BorderColor;
                bottomBorder.Val = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Thick;
                bottomBorder.Size = border.UseVariableBorders ? border.BorderWidthBottom : border.BorderWidth;
                borders.BottomBorder = bottomBorder;
            }
        }
    }
}
