using System.Collections.Generic;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models.Charts
{
    public class BarChartModel
    {
        /// <summary>
        /// Graph Title
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Show / Hide Title
        /// </summary>
        public bool ShowTitle { get; set; }

        /// <summary>
        /// Categories
        /// </summary>
        public IList<string> Categories { get; set; }

        /// <summary>
        /// Taille par défaut des textes dans le graphique
        /// Null par défaut
        /// </summary>
        public double? FontSize { get; set; } = null;

        /// <summary>
        /// Values
        /// </summary>
        public IList<BarSerie> Values { get; set; }

        /// <summary>
        /// Show / Hide Borders
        /// </summary>
        public bool ShowBarBorder { get; set; }
    }
}
