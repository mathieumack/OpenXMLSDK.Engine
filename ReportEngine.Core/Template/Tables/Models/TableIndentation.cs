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
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// default constructor
        /// </summary>
        public TableIndentation()
        {
            Width = 0;
        }
    }
}
