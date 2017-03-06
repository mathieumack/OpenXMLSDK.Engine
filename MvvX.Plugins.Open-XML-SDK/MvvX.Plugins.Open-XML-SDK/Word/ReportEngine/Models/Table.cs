using System.Collections.Generic;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models.Attributes;

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

        /// <summary>
        /// Header row
        /// </summary>
        public Row HeaderRow { get; set; }

        /// <summary>
        /// Footer row
        /// </summary>
        public Row FooterRow { get; set; }

        /// <summary>
        /// Key for datasource
        /// </summary>
        public string DataSourceKey { get; set; }

        /// <summary>
        /// Borders
        /// </summary>
        public BorderModel Borders { get; set; }
    }
}
