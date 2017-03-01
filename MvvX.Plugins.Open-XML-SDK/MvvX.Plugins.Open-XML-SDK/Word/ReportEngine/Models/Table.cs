using System.Collections.Generic;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models
{
    /// <summary>
    /// Model for a table
    /// </summary>
    public class Table : BaseElement
    {
        /// <summary>
        /// Rows of the table
        /// </summary>
        public IList<Row> Rows { get; set; } = new List<Row>();
    }
}
