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
        /// if bind to a datasource, contains the model of a row
        /// </summary>
        public Row RowModel { get; set; }

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

        /// <summary>
        /// Constructor
        /// </summary>
        public Table()
            : base(typeof(Table).Name)
        {
        }
    }
}
