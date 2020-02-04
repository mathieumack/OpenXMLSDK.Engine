using System.Collections.Generic;

namespace OpenXMLSDK.Engine.Word.ReportEngine.BatchModels.Charts
{
    public class SingleSeriesModel
    {
        /// <summary>
        /// Categories
        /// </summary>
        public List<CategoryModel> Categories { get; set; }

        /// <summary>
        /// Value
        /// </summary>
        public SerieModel Serie { get; set; }
    }
}
