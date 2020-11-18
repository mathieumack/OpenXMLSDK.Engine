namespace OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts
{
    /// <summary>
    /// Line graph serie
    /// </summary>
    public class LineSerie : ChartSerie
    {
        /// <summary>
        /// Line Marker
        /// </summary>
        public LineSerieMarker LineSerieMarker { get; set; } = new LineSerieMarker();

        /// <summary>
        /// Default constructor
        /// </summary>
        public LineSerie() { }
    }
}
