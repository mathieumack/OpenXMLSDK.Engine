using System;
using System.Collections.Generic;

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
