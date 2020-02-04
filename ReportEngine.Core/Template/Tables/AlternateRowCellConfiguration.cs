using ReportEngine.Core.Template.ExtendedModels;
using ReportEngine.Core.Template.Tables.Models;

namespace ReportEngine.Core.Template.Tables
{
    public class AlternateRowCellConfiguration
    {
        /// <summary>
        /// Borders
        /// </summary>
        public BorderModel Borders { get; set; }

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
        /// Shading value
        /// </summary>
        public string Shading { get; set; }
    }
}
