namespace MvvX.Plugins.OpenXMLSDK.Word.Tables.Models
{
    public class TableWidthModel
    {
        /// <summary>
        /// Width of the table
        /// default : 5000
        /// </summary>
        public string Width { get; set; }

        /// <summary>
        /// Type of the unit
        /// default : TableWidthUnitValues.Pct
        /// </summary>
        public TableWidthUnitValues Type { get; set; }

        public TableWidthModel()
        {
            Width = "5000";
            Type = TableWidthUnitValues.Pct;
        }
    }
}
