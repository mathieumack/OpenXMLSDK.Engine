using System.Collections.Generic;

namespace MvvX.Plugins.OpenXMLSDK.ReportEngine.BatchModels.Charts
{
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
