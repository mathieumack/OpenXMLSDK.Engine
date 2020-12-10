namespace OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts
{
    /// <summary>
    /// Define Axis' details
    /// </summary>
    public class ChartAxisModel
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
        public string TitleColor { get; set; }

        /// <summary>
        /// Indicate if the axis must be deleted
        /// </summary>
        public bool DeleteAxis { get; set; }

        /// <summary>
        /// Indicate if the curve sepparating axis from graph must be deleted
        /// </summary>
        public bool ShowAxisCurve { get; set; }

        /// <summary>
        /// Define the color of the curve sepparating axis from graph, must be in hex format (with or without #).
        /// Exemple of valid format: 000000 or #000000
        /// </summary>
        public string AxisCurveColor { get; set; }

        /// <summary>
        /// Indicate if we show major grid lines
        /// </summary>
        public bool ShowMajorGridlines { get; set; }

        /// <summary>
        /// Indicate the color of major grid lines, must be in hex format (with or without #).
        /// Exemple of valid format: 000000 or #000000
        /// </summary>
        public string MajorGridlinesColor { get; set; }

        /// <summary>
        /// This element specifies where on the axis the perpendicular axis crosses.
        /// For a category axis, the value is a decimal number on the value axis.
        /// For a value axis, the value is an integer category number, starting with 1 as the first category.
        /// </summary>
        public double? CrossesAt { get; set; }
    }
}
