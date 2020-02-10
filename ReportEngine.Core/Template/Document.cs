using System.Collections.Generic;
using ReportEngine.Core.Template.ExtendedModels;
using ReportEngine.Core.Template.Styles;

namespace ReportEngine.Core.Template
{
    /// <summary>
    /// Model class for a document
    /// </summary>
    public class Document : BaseElement
    {
        /// <summary>
        /// Header of document
        /// </summary>
        public List<Header> Headers { get; set; } = new List<Header>();

        /// <summary>
        /// Footer of document
        /// </summary>
        public List<Footer> Footers { get; set; } = new List<Footer>();

        /// <summary>
        /// List of pages of document
        /// </summary>
        public List<BaseElement> Pages { get; set; } = new List<BaseElement>();

        /// <summary>
        /// List of templates
        /// </summary>
        public List<TemplateDefinition> TemplateDefinitions { get; set; } = new List<TemplateDefinition>();

        /// <summary>
        /// List of styles used in document
        /// </summary>
        public List<Style> Styles { get; set; } = new List<Style>();

        /// <summary>
        /// Margin for all pages of documents
        /// </summary>
        public SpacingModel Margin { get; set; }

        public float DefaultFontSize { get; set; } = 12;

        /// <summary>
        /// Constructor
        /// </summary>
        public Document()
            : base(typeof(Document).Name)
        {
        }
    }
}
