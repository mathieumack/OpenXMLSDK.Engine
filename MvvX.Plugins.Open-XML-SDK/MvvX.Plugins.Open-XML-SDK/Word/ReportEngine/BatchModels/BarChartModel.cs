using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models.Charts;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels
{
    public class BarChartModel : BaseModel
    {
        /// <summary>
        /// Contenu du graphique
        /// </summary>
        public BarChartModel BarChartContent { get; set; }

        public BarChartModel()
            : this(null)
        { }

        public BarChartModel(Models.Charts.BarChartModel barChartContent)
            : base(typeof(BarChartModel).Name)
        {
            this.BarChartContent = barChartContent;
        }
    }
}
