using System.Collections.Generic;
using OpenXMLSDK.Engine.ReportEngine.DataContext.Charts;

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

        /// <summary>
        /// Define the model for categories axis
        /// </summary>
        public AxisModel CategoriesAxisModel { get; set; } = new AxisModel();

        /// <summary>
        /// Define the model for values axis
        /// </summary>
        public AxisModel ValuesAxisModel { get; set; } = new AxisModel();
    }
}
