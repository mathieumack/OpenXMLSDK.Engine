using ReportEngine.Core.Template.Styles;

namespace ReportEngine.Core.Template.Tables.Models
{
    public class TableCellPropertiesModel
    {
        public TableCellWidthModel TableCellWidth { get; set; }

        public ShadingModel Shading { get; set; }

        /// <summary>
        /// Vertical alignement in cell.
        /// Default : Top
        /// </summary>
        public TableVerticalAlignmentValues TableVerticalAlignementValues { get; set; }

        public JustificationValues? Justification { get; set; }

        public TextDirectionValues? TextDirectionValues { get; set; }

        public GridSpanModel Gridspan { get; set; }

        public bool Fusion { get; set; }

        public bool FusionChild { get; set; }

        public string Height { get; set; }

        /// <summary>
        /// Allow to link this cell with the next cell
        /// (For example when you have a new page
        /// Default : False
        /// </summary>
        public bool ParagraphSolidarity { get; set; }

        public TableCellBordersModel TableCellBorders { get; set; }

        public TableCellMarginModel TableCellMargin { get; set; }

        public TableCellPropertiesModel()
        {
            TableVerticalAlignementValues = TableVerticalAlignmentValues.Top;
            ParagraphSolidarity = false;
        }

    }
}
