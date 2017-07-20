using System.Collections.Generic;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models.Charts
{
    public class BarSerie
    {
        /// <summary>
        /// Label
        /// </summary>
        public string Label { get; set; }

        /// <summary>
        /// Values
        /// </summary>
        public IList<double> Values { get; set; }

        /// <summary>
        /// Color (#000000 -> #FFFFFF)
        /// </summary>
        public string RGBbColor { get; set; }

        /// <summary>
        /// Format de rendu des labels
        /// {0} par défaut
        /// </summary>
        public string LabelFormatString { get; set; } = "{0}";
    }
}
