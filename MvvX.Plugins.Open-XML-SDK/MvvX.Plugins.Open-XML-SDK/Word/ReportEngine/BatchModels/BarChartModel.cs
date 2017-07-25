using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels.Charts;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels
{
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
