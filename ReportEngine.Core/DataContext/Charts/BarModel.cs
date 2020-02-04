using System;
using System.Collections.Generic;
using OpenXMLSDK.Engine.Word.ReportEngine.BatchModels.Charts;

namespace ReportEngine.Core.DataContext.Charts
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
