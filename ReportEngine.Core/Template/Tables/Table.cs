using System.Collections.Generic;
using ReportEngine.Core.Template.ExtendedModels;
using ReportEngine.Core.Template.Tables.Models;

namespace ReportEngine.Core.Template.Tables
{
    /// <summary>
    /// Model for a table
    /// </summary>
    public class Table : BaseElement
    {
        /// <summary>
        /// Define the prefix that will be used for automatically added items
        /// IsFirstItem, ...
        /// </summary>
        public string AutoContextAddItemsPrefix { get; set; }

        /// <summary>
        /// Rows of the table, generated before Rows field content
        /// used only if the table is not binded to a datasource (DataSourceKey != null)
        /// </summary>
        public IList<Row> BeforeRows { get; set; } = new List<Row>();

        /// <summary>
        /// Rows of the table
        /// used only if the table is not binded to a datasource (DataSourceKey != null)
        /// </summary>
        public IList<Row> Rows { get; set; } = new List<Row>();

        /// <summary>
        /// Rows of the table, generated after Rows field content
        /// used only if the table is not binded to a datasource (DataSourceKey != null)
        /// </summary>
        public IList<Row> AfterRows { get; set; } = new List<Row>();

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
        /// array containing width of each column 
        /// value is specified in twentieths of a point
        /// these widths determine the initial width of each grid column, which may then be overridden by the table layout algorithm applied to the current table row and the preferred widths of specific cells which are part of that grid column as the table is displayed
        /// </summary>
        public int[] ColsWidth { get; set; }

        /// <summary>
        /// Table Width
        /// value can be in pct (Fiftieths of a Percent) or in dxa (Twentieths of a Point)
        /// </summary>
        public TableWidthModel TableWidth { get; set; }

        /// <summary>
        /// Table indentation
        /// </summary>
        public TableIndentation TableIndentation { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        internal Table(string type)
            : base(type)
        {
            TableIndentation = new TableIndentation();
        }

        /// <summary>
        /// Constructor
        /// </summary>
        public Table()
            : base(typeof(Table).Name)
        {
            TableIndentation = new TableIndentation();
        }
    }
}
