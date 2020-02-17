namespace ReportEngine.Core.Template.Tables.Models
{
    /// <summary>
    /// Table indentation class
    /// </summary>
    public class TableIndentation
    {
        /// <summary>
        /// Table indentation size
        /// default : 0
        /// Defined by default in dxa. <see cref="TableWidthUnitValues"/>
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// Type of the unit
        /// default : TableWidthUnitValues.Pct
        /// </summary>
        public TableWidthUnitValues Type { get; set; } = TableWidthUnitValues.Dxa;
    }
}
