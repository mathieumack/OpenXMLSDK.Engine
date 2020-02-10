namespace ReportEngine.Core.Template.Charts
{
    /// <summary>
    /// Class representating the scaling of a bar chart axis
    /// </summary>
    public class BarChartScalingModel
    {
        /// <summary>
        /// Orientation of the axis
        /// </summary>
        public BarChartOrientationType Orientation { get; set; } = BarChartOrientationType.MinMax;

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