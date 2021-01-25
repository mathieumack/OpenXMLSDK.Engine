using OpenXMLSDK.Engine.Word.ReportEngine.Models.ExtendedModels;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Models
{
    /// <summary>
    /// Model for a page
    /// </summary>
    public class Page : BaseElement
    {
        /// <summary>
        /// Page orientation : Portrait or Landscape
        /// </summary>
        public PageOrientationValues PageOrientation { get; set; }

        /// <summary>
        /// Margin for page
        /// </summary>
        public SpacingModel Margin { get; set; }

        /// <summary>
        /// Number of column for splitting page
        /// </summary>
        public int? ColumnCount { get; set; }

        /// <summary>
        /// Internal constructor used for ForEachPage
        /// </summary>
        internal Page(string type)
            : base(type)
        {
        }

        /// <summary>
        /// Constructor
        /// </summary>
        public Page()
            : base(typeof(Page).Name)
        {
        }
    }
}
