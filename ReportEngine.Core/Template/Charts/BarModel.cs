using System;
using System.Collections.Generic;

namespace ReportEngine.Core.Template.Charts
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
        /// Scaling of the values axis
        /// </summary>
        public BarChartScalingModel ValuesAxisScaling { get; set; }

        /// <summary>
        /// Show / Hide Borders
        /// </summary>
        [Obsolete("Please use ShowChartBorder instead")]
        public bool ShowBarBorder { get; set; }

        /// <summary>
        /// Ctor
        /// </summary>
        public BarModel()
            : base(typeof(BarModel).Name)
        {
            BarChartType = BarChartType.BarChart;
            ValuesAxisScaling = new BarChartScalingModel();
        }
    }
}
