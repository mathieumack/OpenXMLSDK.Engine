namespace OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts
{
    /// <summary>
    /// Marker of line series
    /// </summary>
    public class LineSerieMarker
    {
        /// <summary>
        /// indicate the selected symbol displayed at each point
        /// </summary>
        public MarkerStyleValues MarkerStyleValues { get; set; } = MarkerStyleValues.None;

        /// <summary>
        /// Marker size, MUST be between 2 and 72
        /// </summary>
        public byte Size { get; set; } = 5;

        /// <summary>
        /// Default constructor
        /// </summary>
        public LineSerieMarker() { }
    }
}
