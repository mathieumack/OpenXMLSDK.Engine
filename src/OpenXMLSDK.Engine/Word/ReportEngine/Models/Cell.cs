using OpenXMLSDK.Engine.ReportEngine.DataContext;
using OpenXMLSDK.Engine.Word.ReportEngine.Models.ExtendedModels;
using OpenXMLSDK.Engine.Word.Tables;
using OpenXMLSDK.Engine.Word.Tables.Models;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Models
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
        /// <see cref="DoubleModel"/> linked that will be linked to the ColSpan property
        /// </summary>
        public string ColSpanKey { get; set; }

        /// <summary>
        /// Row Span : Cell is merged vertically : it's the first merged cell
        /// </summary>
        public bool Fusion { get; set; }

        /// <summary>
        /// <see cref="BooleanModel"/> linked that will be linked to the Fusion property
        /// </summary>
        public string FusionKey { get; set; }

        /// <summary>
        /// Cell is a hidden merged part of a rowspan
        /// </summary>
        public bool FusionChild { get; set; }

        /// <summary>
        /// <see cref="BooleanModel"/> linked that will be linked to the FusionChild property
        /// </summary>
        public string FusionChildKey { get; set; }

        /// <summary>
        /// No wrap text 
        /// </summary>
        public bool NoWrap { get; set; }

        /// <summary>
        /// Justification/ horizontal alignment of cells content
        /// </summary>
        public JustificationValues? Justification { get; set; }

        /// <summary>
        /// Vertical alignment of cell content
        /// </summary>
        public TableVerticalAlignmentValues? VerticalAlignment { get; set; }

        /// <summary>
        /// Text direction in cell
        /// </summary>
        public TextDirectionValues? TextDirection { get; set; }

        /// <summary>
        /// Margins in cell
        /// </summary>
        public MarginModel Margin { get; set; }

        /// <summary>
        /// Width of the cell
        /// </summary>
        public TableCellWidthModel CellWidth { get; set; }

        /// <summary>
        /// Alternate configuration is the cell in an alternate row
        /// </summary>
        public AlternateRowCellConfiguration AlternateRowCellConfiguration { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public Cell()
            : base(typeof(Cell).Name)
        {
        }
    }
}
