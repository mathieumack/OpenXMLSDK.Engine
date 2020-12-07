namespace OpenXMLSDK.Engine.ReportEngine.DataContext.Charts
{
    /// <summary>
    /// Define Axis' details
    /// </summary>
    public class AxisModel
    {
        /// <summary>
        /// Title
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Title color, must be in hex format (with or without #).
        /// Exemple of valid format: 000000 or #000000
        /// </summary>
        public string Color { get; set; }
    }
}
