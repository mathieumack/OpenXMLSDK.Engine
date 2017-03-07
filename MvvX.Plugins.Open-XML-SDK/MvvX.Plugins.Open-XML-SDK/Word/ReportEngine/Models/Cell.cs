using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models.Attributes;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models
{
    /// <summary>
    /// Table Cell
    /// </summary>
    public class Cell : BaseElement
    {
        /// <summary>
        /// Borders
        /// </summary>
        public BorderModel Borders { get; set; }

        /// <summary>
        /// Colspan
        /// </summary>
        public int ColSpan { get; set; }

        /// <summary>
        /// Row Span : Cell is merged vertically : it's the first merged cell
        /// </summary>
        public bool Fusion { get; set; }

        /// <summary>
        /// Cell is a hidden merged part of a rowspan
        /// </summary>
        public bool FusionChild { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public Cell()
            : base(typeof(Cell).Name)
        {
        }
    }
}
