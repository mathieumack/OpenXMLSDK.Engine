using System.Collections.Generic;
using OpenXMLSDK.Engine.Word.ReportEngine.BatchModels.Charts;

namespace OpenXMLSDK.Engine.ReportEngine.DataContext.Charts
{
    /// <summary>
    /// Scatter serie
    /// </summary>
    public class ScatterSerieModel : SerieModel
    {
        /// <summary>
        /// Coordinate of each value
        /// </summary>
        public new List<CurvePoint> Values { get; set; }

        /// <summary>
        /// Indicate if the curve must be hidden
        /// </summary>
        public bool? HideCurve { get; set; }
    }
}
