namespace ReportEngine.Core.Template.ExtendedModels
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

        /// <summary>
        /// Distance (in twentieths of a point) from the top edge of the page to the top edge of the header.
        /// </summary>
        public uint Header { get; set; }

        /// <summary>
        /// Distance (in twentieths of a point) from the bottom edge of the page to the bottom edge of the footer.
        /// </summary>
        public uint Footer { get; set; }

    }
}
