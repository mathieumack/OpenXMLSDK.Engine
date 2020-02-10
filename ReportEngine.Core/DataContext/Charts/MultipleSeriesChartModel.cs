
namespace ReportEngine.Core.DataContext.Charts
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
