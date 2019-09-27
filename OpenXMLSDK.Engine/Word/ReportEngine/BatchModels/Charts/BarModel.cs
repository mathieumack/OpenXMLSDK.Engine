using System;
using System.Collections.Generic;

namespace OpenXMLSDK.Engine.Word.ReportEngine.BatchModels.Charts
{
    [Obsolete("Please use MultipleSeriesModel instead")]
    public class BarModel
    {
        /// <summary>
        /// Categories
        /// </summary>
        public List<BarCategoryModel> Categories { get; set; }
        
        /// <summary>
        /// Values
        /// </summary>
        public List<BarSerieModel> Series { get; set; }
    }
}
