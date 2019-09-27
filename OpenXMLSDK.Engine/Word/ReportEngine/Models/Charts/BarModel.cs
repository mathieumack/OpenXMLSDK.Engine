using System.Collections.Generic;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts
{
    public class BarModel : ChartModel
    {
        /// <summary>
        /// Type of the barchart
        /// </summary>
        public BarChartType BarChartType { get; set; }

        /// <summary>
        /// Direction of bar chart
        /// Horizontal = Bar chart (default)
        /// Vertical = Column chart
        /// </summary>
        public BarDirectionValues BarDirectionValues { get; set; } = BarDirectionValues.Bar;

        /// <summary>
        /// Type of bar grouping
        /// </summary>
        public BarGroupingValues BarGroupingValues { get; set; } = BarGroupingValues.Stacked;

        /// <summary>
        /// Categories
        /// </summary>
        public List<BarCategory> Categories { get; set; }

        /// <summary>
        /// Values
        /// </summary>
        public List<BarSerie> Series { get; set; }

        /// <summary>
        /// Ctor
        /// </summary>
        public BarModel()
            : base(typeof(BarModel).Name)
        {
            BarChartType = BarChartType.BarChart;
        }
    }
}
