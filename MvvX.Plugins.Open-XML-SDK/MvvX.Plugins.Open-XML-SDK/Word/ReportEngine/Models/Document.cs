using System.Collections.Generic;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models
{
    /// <summary>
    /// Model class for a document
    /// </summary>
    public class Document : BaseElement
    {
        /// <summary>
        /// List of pages of document
        /// </summary>
        public IList<Page> Pages { get; set; } = new List<Page>();

        /// <summary>
        /// List of styles used in document
        /// </summary>
        public IList<Style> Styles { get; set; } = new List<Style>();
    }
}
