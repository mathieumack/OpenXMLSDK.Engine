namespace MvvX.Open_XML_SDK.Core.Word.Tables.Models
{
    public class TableCellPropertiesModel
    {
        public string Width { get; set; }

        public TableWidthUnitValue WidthUnit { get; set; }

        //TODO
        //public Shading Shading { get; set; }

        public TableVerticalAlignmentValue TableVerticalAlignementValues { get; set; }

        public JustificationValue? Justification { get; set; }

        public TextDirectionValue? TextDirectionValues { get; set; }

        public int? Gridspan { get; set; }

        public bool Fusion { get; set; }

        public bool FusionChild { get; set; }

        public string Height { get; set; }

        public TableCellPropertiesModel()
        {
            TableVerticalAlignementValues = TableVerticalAlignmentValue.Top;
        }

    }
}
