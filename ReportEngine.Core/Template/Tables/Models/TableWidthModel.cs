namespace ReportEngine.Core.Template.Tables.Models
{
    /// <summary>
    /// Table width class
    /// </summary>
    public class TableWidthModel
    {
        /// <summary>
        /// Width of the table
        /// default : 5000
        /// </summary>
        public float Width { get; set; } = 5000;

        /// <summary>
        /// Type of the unit
        /// default : TableWidthUnitValues.Pct
        /// </summary>
        public TableWidthUnitValues Type { get; set; } = TableWidthUnitValues.Pct;
    }
}
