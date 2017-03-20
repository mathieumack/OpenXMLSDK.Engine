using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine
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
    }
}
