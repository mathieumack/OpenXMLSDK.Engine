namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models.Attributes
{
    /// <summary>
    /// Model class for margin
    /// </summary>
    public class SpacingModel
    {
        /// <summary>
        /// Distance (in twentieths of a point) between the left edge of the page and the left edge of the text extents for this document.
        /// </summary>
        public uint Left { get; set; }

        /// <summary>
        /// Distance (in twentieths of a point) between the top of the text margins for the main document and the top of the page
        /// </summary>
        public int Top { get; set; }

        /// <summary>
        /// Distance (in twentieths of a point) between the right edge of the page and the right edge of the text extents for this document
        /// </summary>
        public uint Right { get; set; }

        /// <summary>
        /// Distance (in twentieths of a point) between the bottom of the text margins for the main document and the bottom of the page
        /// </summary>
        public int Bottom { get; set; }
    }
}
