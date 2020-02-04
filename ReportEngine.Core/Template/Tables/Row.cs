using System.Collections.Generic;

namespace ReportEngine.Core.Template.Tables
{
    public class Row : BaseElement
    {
        /// <summary>
        /// Cells of the row
        /// </summary>
        public IList<Cell> Cells { get; set; } = new List<Cell>();

        /// <summary>
        /// Table Width
        /// Row height
        /// </summary>
        public int? RowHeight { get; set; }

        /// <summary>
        /// CantSplit
        /// </summary>
        public bool CantSplit { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public Row()
            : base(typeof(Row).Name)
        {
        }
    }
}
