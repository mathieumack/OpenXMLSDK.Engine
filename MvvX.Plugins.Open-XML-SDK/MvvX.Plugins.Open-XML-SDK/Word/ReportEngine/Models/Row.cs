using System.Collections.Generic;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models
{
    public class Row : BaseElement
    {
        /// <summary>
        /// Cells of the row
        /// </summary>
        public IList<Cell> Cells { get; set; } = new List<Cell>();
    }
}
