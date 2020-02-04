using ReportEngine.Core.Template.ExtendedModels;

namespace ReportEngine.Core.Template
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
