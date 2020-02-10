using System.Collections.Generic;

namespace ReportEngine.Core.DataContext.Charts
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
