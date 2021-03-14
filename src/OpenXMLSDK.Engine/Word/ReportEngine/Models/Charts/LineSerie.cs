using OpenXMLSDK.Engine.ReportEngine.DataContext.Charts;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts
{
    /// <summary>
    /// Line graph serie
    /// </summary>
    public class LineSerie : ChartSerie
    {
        /// <summary>
        /// Serie Marker
        /// </summary>
        public SerieMarker LineSerieMarker { get; set; } = new SerieMarker();

        /// <summary>
        /// Default constructor
        /// </summary>
        public LineSerie() { }
    }
}
