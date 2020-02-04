using System.Collections.Generic;

namespace ReportEngine.Core.Template.Text
{
    /// <summary>
    /// Model for an hyperlink
    /// Anchor > internal link to a bookmark
    /// WebSiteUri > external web site uri
    /// </summary>
    public class Hyperlink : BaseElement
    {
        /// <summary>
        /// Web site uri
        /// </summary>
        public string WebSiteUri { get; set; }

        /// <summary>
        /// Anchor label, can be contextualized
        /// </summary>
        public string Anchor { get; set; }

        /// <summary>
        /// Content text
        /// </summary>
        public Label Text { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public Hyperlink()
            : base(typeof(Hyperlink).Name)
        {
            Anchor = string.Empty;
        }
    }
}
