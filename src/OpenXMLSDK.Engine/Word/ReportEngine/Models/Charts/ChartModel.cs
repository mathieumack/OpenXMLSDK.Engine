using System;
using OpenXMLSDK.Engine.ReportEngine.DataContext.Charts;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts
{
    public class ChartModel : BaseElement
    {
        /// <summary>
        /// Chart data source
        /// </summary>
        public string DataSourceKey { get; set; }

        /// <summary>
        /// Graph Title
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Show / Hide Title
        /// </summary>
        public bool ShowTitle { get; set; }

        /// <summary>
        /// Indicate if we delete axis for values
        /// </summary>
        [Obsolete("Please use ValuesAxisModel.DeleteAxis instead")]
        public bool DeleteAxeValue { get; set; }

        /// <summary>
        /// Space between line categories
        /// </summary>
        public int? SpaceBetweenLineCategories { get; set; }

        /// <summary>
        /// Indicate if we delete axis for categories
        /// </summary>
        [Obsolete("Please use CategoriesAxisModel.DeleteAxis instead")]
        public bool DeleteAxeCategory { get; set; }

        /// <summary>
        /// Indicate if the legend must be displayed
        /// </summary>
        public bool ShowLegend { get; set; }

        /// <summary>
        /// Legend font family
        /// </summary>
        public string FontFamilyLegend { get; set; }

        /// <summary>
        /// Specify the legend position
        /// </summary>
        public LegendPositionValues LegendPosition { get; set; } = LegendPositionValues.Right;

        /// <summary>
        /// Define if the graph has a border
        /// </summary>
        public bool HasBorder { get; set; }

        /// <summary>
        /// Max width of the graph
        /// </summary>
        public double? MaxWidth { get; set; }

        /// <summary>
        /// Max height of the graph
        /// </summary>
        public double? MaxHeight { get; set; }

        /// <summary>
        /// Show / Hide Borders
        /// </summary>
        public bool ShowChartBorder { get; set; }

        /// <summary>
        /// Indicate if we show major grid lines
        /// </summary>
        [Obsolete("Please use CategoriesAxisModel.ShowMajorGridlines instead")]
        public bool ShowMajorGridlines { get; set; }

        /// <summary>
        /// Indicate the color of major grid lines
        /// </summary>
        [Obsolete("Please use CategoriesAxisModel.MajorGridlinesColor instead")]
        public string MajorGridlinesColor { get; set; }

        /// <summary>
        /// Border color
        /// </summary>
        public string BorderColor { get; set; }

        /// <summary>
        /// Border width
        /// </summary>
        public int? BorderWidth { get; set; }

        /// <summary>
        /// Rounded corner for border
        /// </summary>
        public bool RoundedCorner { get; set; }

        /// <summary>
        /// Color of data labels
        /// </summary>
        public string DataLabelColor { get; set; }

        /// <summary>
        /// Options for data label
        /// </summary>
        public DataLabelModel DataLabel { get; set; }

        [Obsolete("Please use DataLabel.ShowDataLabel instead")]
        public bool ShowDataLabel
        {
            get
            {
                if (DataLabel == null)
                    return true;
                else
                    return DataLabel.ShowDataLabel;
            }
            set
            {
                if (DataLabel == null)
                    DataLabel = new DataLabelModel();
                DataLabel.ShowDataLabel = value;
            }
        }

        /// <summary>
        /// Specifies that each data marker in the series has a different color
        /// </summary>
        public bool VaryColors { get; set; } = true;

        /// <summary>
        /// Define the model for categories axis
        /// </summary>
        public ChartAxisModel CategoriesAxisModel { get; set; } = new ChartAxisModel
        {
            TickLabelPosition = TickLabelPositionValues.Low
        };

        /// <summary>
        /// Define the model for values axis
        /// </summary>
        public ChartAxisModel ValuesAxisModel { get; set; } = new ChartAxisModel
        {
            TickLabelPosition = TickLabelPositionValues.NextTo
        };

        /// <summary>
        /// Define the model for secondary categories axis
        /// </summary>
        public ChartAxisModel SecondaryCategoriesAxisModel { get; set; } = new ChartAxisModel
        {
            TickLabelPosition = TickLabelPositionValues.High
        };

        /// <summary>
        /// Define the model for secondary values axis
        /// </summary>
        public ChartAxisModel SecondaryValuesAxisModel { get; set; } = new ChartAxisModel
        {
            TickLabelPosition = TickLabelPositionValues.NextTo
        };

        /// <summary>
        /// Indicate the overlap of bar series in percent, Min = -100, Max = 100
        /// </summary>
        public sbyte Overlap { get; set; } = 100;

        /// <summary>
        /// Indicate the category value type
        /// </summary>
        public CategoryType CategoryType { get; set; }

        /// <summary>
        /// Ctor
        /// </summary>
        public ChartModel(string type) : base(type)
        {
        }
    }
}
