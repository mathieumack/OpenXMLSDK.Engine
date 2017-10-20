using MvvX.Plugins.OpenXMLSDK.Word.Tables.Models;
using System.Collections.Generic;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models
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
        /// Indicates whether row can be split or not.
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
