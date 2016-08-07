namespace MvvX.Open_XML_SDK.Core.Word.Tables.Models
{
    public class TableCellPropertiesModel
    {
        /// <summary>
        /// Width of the cell
        /// </summary>
        public string Width { get; set;}

        public TableWidthUnitValues WidthUnit { get; set; }

        public ShadingPatternValues? Shading { get; set; }

        /// <summary>
        /// Vertical alignement in cell.
        /// Default : Top
        /// </summary>
        public TableVerticalAlignmentValues TableVerticalAlignementValues { get; set; }

        public JustificationValues? Justification { get; set; }

        public TextDirectionValues? TextDirectionValues { get; set; }

        public int? Gridspan { get; set; }

        public bool Fusion { get; set; }

        public bool FusionChild { get; set; }

        public string Height { get; set; }

        /// <summary>
        /// Allow to link this cell with the next cell
        /// (For example when you have a new page
        /// Default : False
        /// </summary>
        public bool ParagraphSolidarity { get; set; }

        /// <summary>
        /// Type of the top border.
        /// Default : null, not set.
        /// </summary>
        public TableBorderModel TopBorder { get; set; }

        /// <summary>
        /// Type of the bottom border.
        /// Default : null, not set.
        /// </summary>
        public TableBorderModel BottomBorder { get; set; }

        /// <summary>
        /// Type of the left border.
        /// Default : null, not set.
        /// </summary>
        public TableBorderModel LeftBorder { get; set; }

        /// <summary>
        /// Type of the right border.
        /// Default : null, not set.
        /// </summary>
        public TableBorderModel RightBorder { get; set; }

        public TableCellPropertiesModel()
        {
            TableVerticalAlignementValues = TableVerticalAlignmentValues.Top;
            ParagraphSolidarity = false;
        }

    }
}
