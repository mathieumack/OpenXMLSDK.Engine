using System.Collections.Generic;

namespace ReportEngine.Core.Template.Charts
{
    public class PieModel : ChartModel
    {
        /// <summary>
        /// Type of the pieChart
        /// </summary>
        public PieChartType PieChartType { get; set; }

        /// <summary>
        /// Categories
        /// </summary>
        public List<PieCategory> Categories { get; set; }

        /// <summary>
        /// Values
        /// </summary>
        public PieSerie Serie { get; set; }

        /// <summary>
        /// Ctor
        /// </summary>
        public PieModel()
            : base(typeof(PieModel).Name)
        {
        }
    }
}
