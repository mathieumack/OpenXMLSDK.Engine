namespace OpenXMLSDK.Engine.ReportEngine.DataContext.Charts
{
    /// <summary>
    /// Series' marker
    /// </summary>
    public class SerieMarker
    {
        /// <summary>
        /// Indicate the selected symbol displayed at each point.
        /// </summary>
        public MarkerStyleValues MarkerStyleValues { get; set; } = MarkerStyleValues.None;

        /// <summary>
        /// Marker size, MUST be between 2 and 72.
        /// Default value is 5.
        /// </summary>
        public byte Size { get; set; } = 5;
    }
}
