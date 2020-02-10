using ReportEngine.Core.Template.ExtendedModels;

namespace ReportEngine.Core.Template.Text
{
    /// <summary>
    /// Model for paragraph
    /// </summary>
    public class Paragraph : BaseElement
    {
        /// <summary>
        /// Justification of text inside paragraph
        /// </summary>
        public JustificationValues Justification { get; set; }

        /// <summary>
        /// Spacing above the first line in this paragraph, in twentieths of a point
        /// </summary>
        public int? SpacingBefore { get; set; }

        /// <summary>
        /// Spacing after the last line, in twentieths of a point
        /// </summary>
        public int? SpacingAfter { get; set; }

        /// <summary>
        /// Spacing between lines of text within paragraph, in 240ths of line
        /// </summary>
        public int? SpacingBetweenLines { get; set; }

        /// <summary>
        /// Id of style
        /// </summary>
        public string ParagraphStyleId { get; set; }

        /// <summary>
        /// Borders
        /// </summary>
        public BorderModel Borders { get; set; }

        /// <summary>
        /// Keeplines
        /// </summary>
        public bool Keeplines { get; set; }

        /// <summary>
        /// KeepNext
        /// </summary>
        public bool KeepNext { get; set; }

        /// <summary>
        /// Indicate if a page break must be include before the Paragraph
        /// https://docs.microsoft.com/fr-fr/dotnet/api/documentformat.openxml.wordprocessing.pagebreakbefore?view=openxml-2.8.1
        /// </summary>
        public bool PageBreakBefore { get; set; }

        /// <summary>
        /// Change the PageBreakBefore value from a ContextModel key
        /// </summary>
        public string PageBreakBeforeKey { get; set; }
        
        /// <summary>
        /// Identation properties
        /// </summary>
        public ParagraphIndentationModel Indentation { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public Paragraph()
            : base(typeof(Paragraph).Name)
        {
        }
    }
}
