namespace MvvX.Open_XML_SDK.Core.Word.Tables.Models
{
    public class TablePropertiesModel
    {
        public TableBorderModel TopBorder { get; set; }

        public TableBorderModel BottomBorder { get; set; }

        public TableBorderModel LeftBorder { get; set; }

        public TableBorderModel RightBorder { get; set; }

        public TableBorderModel InsideHorizontalBorder { get; set; }

        public TableBorderModel InsideVerticalBorder { get; set; }

        public TableRowAlignmentValue TableJustification { get; set; }

        public string Width { get; set; }

        public TableWidthUnitValue WidthUnit { get; set; }

        public TablePropertiesModel()
        {
            TopBorder = new TableBorderModel();
            BottomBorder = new TableBorderModel();
            LeftBorder = new TableBorderModel();
            RightBorder = new TableBorderModel();
            InsideHorizontalBorder = new TableBorderModel();
            InsideVerticalBorder = new TableBorderModel();
            Width = "5000";
            WidthUnit = TableWidthUnitValue.Pct;
            TableJustification = TableRowAlignmentValue.Center;
        }
    }
}
