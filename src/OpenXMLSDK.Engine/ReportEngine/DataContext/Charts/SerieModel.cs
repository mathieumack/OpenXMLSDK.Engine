using System.Collections.Generic;

namespace OpenXMLSDK.Engine.ReportEngine.DataContext.Charts
{
    public class SerieModel
    {
        /// <summary>
        /// Name of the serie
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Values
        /// </summary>
        public List<double?> Values { get; set; }

        /// <summary>
        /// Color of the serie
        /// </summary>
        public string Color { get; set; }

        /// <summary>
        /// Color of labels
        /// </summary>
        public string DataLabelColor { get; set; }

        /// <summary>
        /// Labels rendering format
        /// {0} by default
        /// </summary>
        public string LabelFormatString { get; set; } = "{0}";

        /// <summary>
        /// Define if the serie has a border
        /// </summary>
        public bool HasBorder { get; set; }

        /// <summary>
        /// Border color
        /// </summary>
        public string BorderColor { get; set; }

        /// <summary>
        /// Border width
        /// </summary>
        public int? BorderWidth { get; set; }

        /// <summary>
        /// Indicate if this serie is on the first or the secondary axis
        /// </summary>
        public bool UseSecondaryAxis { get; set; }

        /// <summary>
        /// Indicate if the curve must be smooth
        /// </summary>
        public bool SmoothCurve { get; set; }

        /// <summary>
        /// Indicate the curve line style
        /// </summary>
        public PresetLineDashValues PresetLineDashValues { get; set; }

        /// <summary>
        /// Sepcify the render type of the serie (line, bar, ...)
        /// </summary>
        public SerieChartType SerieChartType { get; set; }

        /// <summary>
        /// Specify the marker used at each point (none by default)
        /// </summary>
        public SerieMarker SerieMarker { get; set; } = new SerieMarker();
    }
}
