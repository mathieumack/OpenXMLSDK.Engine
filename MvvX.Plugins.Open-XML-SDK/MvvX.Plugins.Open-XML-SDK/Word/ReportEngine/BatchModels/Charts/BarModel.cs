using System.Collections.Generic;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels.Charts
{
    public class BarModel
    {
        /// <summary>
        /// Graph Title
        /// </summary>
        public string Title { get; set; }
        
        /// <summary>
        /// Categories
        /// </summary>
        public List<BarCategoryModel> Categories { get; set; }
        
        /// <summary>
        /// Color of data labels
        /// </summary>
        public string DataLabelColor { get; set; }

        /// <summary>
        /// Values
        /// </summary>
        public List<BarSerieModel> Series { get; set; }
    }
}
