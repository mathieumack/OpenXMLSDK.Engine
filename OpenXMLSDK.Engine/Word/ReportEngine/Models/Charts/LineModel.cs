using System.Collections.Generic;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts
{
    /// <summary>
    /// Line graph model
    /// </summary>
    public class LineModel : ChartModel
    {
        /// <summary>
        /// Indicate the wanted curve group, by default curves will not be grouped
        /// </summary>
        public GroupingValues GroupingValues { get; set; } = GroupingValues.Standard;

        /// <summary>
        /// Categories
        /// </summary>
        public List<LineCategory> Categories { get; set; }

        /// <summary>
        /// Series (values)
        /// </summary>
        public List<LineSerie> Series { get; set; }

        /// <summary>
        /// Values' axis scaling
        /// </summary>
        public ScalingModel ValuesAxisScaling { get; set; } = new ScalingModel();

        /// <summary>
        /// Constructor
        /// </summary>
        public LineModel() : base(typeof(LineModel).Name)
        {
            VaryColors = false;
        }
    }
}
