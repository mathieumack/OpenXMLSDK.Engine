using System;
using OpenXMLSDK.Engine.Word.ReportEngine.BatchModels.Charts;

namespace OpenXMLSDK.Engine.Word.ReportEngine.BatchModels
{
    [Obsolete("Please use MultipleSeriesChartModel instead")]
    public class BarChartModel : BaseModel
    {
        /// <summary>
        /// Contenu du graphique
        /// </summary>
        public BarModel BarChartContent { get; set; }

        public BarChartModel()
            : this(null)
        { }

        public BarChartModel(BarModel barChartContent)
            : base(typeof(BarChartModel).Name)
        {
            this.BarChartContent = barChartContent;
        }
    }
}
