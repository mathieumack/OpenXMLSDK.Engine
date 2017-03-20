using System.Collections.Generic;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models.Attributes;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models
{
    /// <summary>
    /// Model class for a document
    /// </summary>
    public class Document : BaseElement
    {
        /// <summary>
        /// Header of document
        /// </summary>
        public IList<Header> Headers { get; set; } = new List<Header>();

        /// <summary>
        /// Footer of document
        /// </summary>
        public IList<Footer> Footers { get; set; } = new List<Footer>();

        /// <summary>
        /// List of pages of document
        /// </summary>
        public IList<Page> Pages { get; set; } = new List<Page>();

        /// <summary>
        /// List of styles used in document
        /// </summary>
        public IList<Style> Styles { get; set; } = new List<Style>();

        /// <summary>
        /// Margin for all pages of documents
        /// </summary>
        public SpacingModel Margin { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public Document()
            : base(typeof(Document).Name)
        {
        }
    }
}
