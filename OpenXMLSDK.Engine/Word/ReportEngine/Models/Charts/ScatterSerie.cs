using System.Collections.Generic;
using OpenXMLSDK.Engine.ReportEngine.DataContext.Charts;
using OpenXMLSDK.Engine.Word.ReportEngine.BatchModels.Charts;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts
{
    /// <summary>
    /// Scatter graph serie
    /// </summary>
    public class ScatterSerie : ChartSerie
    {
        /// <summary>
        /// Coordinate of each value
        /// </summary>
        public new List<CurvePoint> Values { get; set; }

        /// <summary>
        /// Indicate if the curve must be hidden
        /// </summary>
        public bool? HideCurve { get; set; }

        /// <summary>
        /// Serie Marker
        /// </summary>
        public SerieMarker SerieMarker { get; set; } = new SerieMarker();
    }
}
