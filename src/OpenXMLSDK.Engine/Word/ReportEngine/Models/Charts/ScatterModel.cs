using System.Collections.Generic;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts
{
    public class ScatterModel : ChartModel
    {
        /// <summary>
        /// Series
        /// </summary>
        public List<ScatterSerie> Series { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public ScatterModel() : base(typeof(ScatterModel).Name)
        {
            VaryColors = false;
        }
    }
}
