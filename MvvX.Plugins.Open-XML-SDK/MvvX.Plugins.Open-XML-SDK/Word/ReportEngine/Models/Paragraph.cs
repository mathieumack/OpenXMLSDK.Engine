using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models.Attributes;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models
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
        /// Constructor
        /// </summary>
        public Paragraph()
            : base(typeof(Paragraph).Name)
        {
        }
    }
}
