using System;

namespace ReportEngine.Core.Template.Charts
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
        /// Indicate if we show major grid lines
        /// </summary>
        public bool ShowLegend { get; set; }

        /// <summary>
        /// Indicate if we delete axis for values
        /// </summary>
        public bool DeleteAxeValue { get; set; }

        /// <summary>
        /// Space between line categories
        /// </summary>
        public int? SpaceBetweenLineCategories { get; set; }

        /// <summary>
        /// Indicate if we delete axis for categories
        /// </summary>
        public bool DeleteAxeCategory { get; set; }

        /// <summary>
        /// Legend font family
        /// </summary>
        public string FontFamilyLegend { get; set; }

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
        public bool ShowMajorGridlines { get; set; }

        /// <summary>
        /// Indicate the color of major grid lines
        /// </summary>
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
        /// Ctor
        /// </summary>
        public ChartModel(string type) : base(type)
        {
        }
    }
}
