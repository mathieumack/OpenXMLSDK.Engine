using OpenXMLSDK.Engine.Word.ReportEngine.BatchModels.Charts;

namespace OpenXMLSDK.Engine.Word.ReportEngine.BatchModels
{
    public class SingleSerieChartModel : BaseModel
    {
        /// <summary>
        /// Contenu du graphique
        /// </summary>
        public SingleSeriesModel ChartContent { get; set; }

        public SingleSerieChartModel()
            : this(null)
        { }

        public SingleSerieChartModel(SingleSeriesModel chartContent)
            : base(typeof(SingleSerieChartModel).Name)
        {
            this.ChartContent = chartContent;
        }
    }
}
