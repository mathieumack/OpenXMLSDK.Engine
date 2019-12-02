using OpenXMLSDK.Engine.ReportEngine.DataContext;
using OpenXMLSDK.Engine.Word.ReportEngine.Models;

namespace OpenXMLSDK.Engine.Word.ReportEngine
{
    /// <summary>
    /// Report serialization class
    /// </summary>
    public class Report
    {
        /// <summary>
        /// Document / Template
        /// </summary>
        public Document Document { get; set; }

        /// <summary>
        /// Context
        /// </summary>
        public ContextModel ContextModel { get; set; }

        /// <summary>
        /// Indicates whether or not a page break is added at the end of report.
        /// </summary>
        public bool AddPageBreak { get; set; }
    }
}
