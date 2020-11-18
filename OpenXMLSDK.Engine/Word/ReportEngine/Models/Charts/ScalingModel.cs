namespace OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts
{
    /// <summary>
    /// Graph axis scaling model
    /// </summary>
    public class ScalingModel
    {
        /// <summary>
        /// Orientation of the axis
        /// </summary>
        public OrientationType Orientation { get; set; } = OrientationType.MinMax;

        /// <summary>
        /// Minimum axis value
        /// </summary>
        public double? MinAxisValue { get; set; }

        /// <summary>
        /// Maximum axis value
        /// </summary>
        public double? MaxAxisValue { get; set; }
    }
}
