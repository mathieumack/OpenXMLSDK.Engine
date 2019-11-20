using OpenXMLSDK.Engine.Word.ReportEngine.BatchModels.Charts;

namespace OpenXMLSDK.Engine.ReportEngine.DataContext
{
    public class MultipleSeriesChartModel : BaseModel
    {
        /// <summary>
        /// Contenu du graphique
        /// </summary>
        public MultipleSeriesModel ChartContent { get; set; }

        public MultipleSeriesChartModel()
            : this(null)
        { }

        public MultipleSeriesChartModel(MultipleSeriesModel chartContent)
            : base(typeof(MultipleSeriesChartModel).Name)
        {
            this.ChartContent = chartContent;
        }
    }
}
