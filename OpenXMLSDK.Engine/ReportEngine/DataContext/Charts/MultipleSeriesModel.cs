using System.Collections.Generic;
using OpenXMLSDK.Engine.ReportEngine.DataContext.Charts;

namespace OpenXMLSDK.Engine.Word.ReportEngine.BatchModels.Charts
{
    public class MultipleSeriesModel
    {
        /// <summary>
        /// Indicate the category value type
        /// </summary>
        public CategoryTypes? CategoryType { get; set; }

        /// <summary>
        /// Categories
        /// </summary>
        public List<CategoryModel> Categories { get; set; }

        /// <summary>
        /// Values
        /// </summary>
        public List<SerieModel> Series { get; set; }

        /// <summary>
        /// Series for a Scatter graph
        /// </summary>
        public List<ScatterSerieModel> ScatterSeries { get; set; }

        /// <summary>
        /// Define the model for categories axis
        /// </summary>
        public AxisModel CategoriesAxisModel { get; set; } = new AxisModel();

        /// <summary>
        /// Define the model for values axis
        /// </summary>
        public AxisModel ValuesAxisModel { get; set; } = new AxisModel();

        /// <summary>
        /// Define the model for secondary categories axis
        /// </summary>
        public AxisModel SecondaryCategoriesAxisModel { get; set; } = new AxisModel();

        /// <summary>
        /// Define the model for secondary values axis
        /// </summary>
        public AxisModel SecondaryValuesAxisModel { get; set; } = new AxisModel();
    }
}
