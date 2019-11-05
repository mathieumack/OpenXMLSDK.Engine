using System.Collections.Generic;

namespace OpenXMLSDK.Engine.Word.ReportEngine.BatchModels.Charts
{
    public class MultipleSeriesModel
    {
        /// <summary>
        /// Categories
        /// </summary>
        public List<CategoryModel> Categories { get; set; }

        /// <summary>
        /// Values
        /// </summary>
        public List<SerieModel> Series { get; set; }
    }
}
