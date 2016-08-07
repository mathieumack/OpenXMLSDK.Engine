namespace MvvX.Open_XML_SDK.Core.Word.Tables.Models
{
    public class TablePropertiesModel
    {
        public TableBordersModel TableBorders { get; set; }

        /// <summary>
        /// Justification of the table
        /// default : Center
        /// </summary>
        public TableRowAlignmentValues TableJustification { get; set; }

        public string Width { get; set; }

        public TableWidthUnitValues WidthUnit { get; set; }

        public TablePropertiesModel()
        {
            Width = "5000";
            WidthUnit = TableWidthUnitValues.Pct;
            TableJustification = TableRowAlignmentValues.Center;
        }
    }
}
