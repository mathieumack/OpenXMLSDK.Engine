namespace ReportEngine.Core.Template.Tables.Models
{
    public class TableCellWidthModel
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
    }
}
