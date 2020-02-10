using System.Collections.Generic;

namespace ReportEngine.Core.DataContext.Charts
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
