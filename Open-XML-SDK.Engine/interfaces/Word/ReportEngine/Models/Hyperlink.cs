using System.Collections.Generic;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Models
{
    /// <summary>
    /// Model for an hyperlink
    /// </summary>
    public class Hyperlink : BaseElement
    {
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
