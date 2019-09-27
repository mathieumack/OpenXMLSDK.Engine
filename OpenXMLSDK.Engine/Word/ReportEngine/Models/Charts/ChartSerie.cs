using System.Collections.Generic;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts
{
    public class ChartSerie : BaseElement
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

        public ChartSerie()
            : base(typeof(ChartSerie).Name)
        {
        }
    }
}
