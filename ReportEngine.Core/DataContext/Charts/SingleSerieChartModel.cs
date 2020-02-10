namespace ReportEngine.Core.DataContext.Charts
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
