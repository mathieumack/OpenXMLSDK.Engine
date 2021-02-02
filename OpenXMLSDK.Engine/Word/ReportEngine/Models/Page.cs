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
        /// Column for page
        /// </summary>
        public Column Column { get; set; }

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
