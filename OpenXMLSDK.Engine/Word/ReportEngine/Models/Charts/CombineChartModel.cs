using System.Collections.Generic;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts
{
    public class CombineChartModel : ChartModel
    {
        /// <summary>
        /// Indicate the wanted curve group, by default curves will not be grouped
        /// </summary>
        public GroupingValues GroupingValues { get; set; } = GroupingValues.Standard;

        /// <summary>
        /// Categories
        /// </summary>
        public List<ChartCategory> Categories { get; set; }

        /// <summary>
        /// line Series (values)
        /// </summary>
        public List<LineSerie> LineSeries { get; set; }

        /// <summary>
        /// Bar Series (values)
        /// </summary>
        public List<BarSerie> BarSeries { get; set; }

        /// <summary>
        /// Values' axis scaling
        /// </summary>
        public ScalingModel ValuesAxisScaling { get; set; } = new ScalingModel();

        /// <summary>
        /// Default constructor
        /// </summary>
        public CombineChartModel() : base(typeof(CombineChartModel).Name)
        {

        }
    }
}
