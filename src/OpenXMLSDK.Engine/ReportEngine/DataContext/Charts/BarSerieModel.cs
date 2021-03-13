using System;
using System.Collections.Generic;

namespace OpenXMLSDK.Engine.ReportEngine.DataContext.Charts
{
    [Obsolete("Please use SerieModel instead")]
    public class BarSerieModel
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
        /// Format de rendu des labels
        /// {0} par défaut
        /// </summary>
        public string LabelFormatString { get; set; } = "{0}";
    }
}
