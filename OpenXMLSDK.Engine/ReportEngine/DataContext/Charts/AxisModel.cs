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
        /// Specify the dispay format
        /// </summary>
        public string LabelFormat { get; set; } = "{0}";

        /// <summary>
        /// Title color, must be in hex format (with or without #).
        /// Exemple of valid format: 000000 or #000000
        /// </summary>
        public string Color { get; set; }

        /// <summary>
        /// This element specifies where on the axis the perpendicular axis crosses.
        /// For a category axis, the value is a decimal number on the value axis.
        /// For a value axis, the value is an integer category number, starting with 1 as the first category.
        /// </summary>
        public double? CrossesAt { get; set; }
    }
}
