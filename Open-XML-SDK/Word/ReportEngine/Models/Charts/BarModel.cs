using System.Collections.Generic;

namespace OpenXMLSDK.Word.ReportEngine.Models.Charts
{
    public class BarModel : BaseElement
    {
        /// <summary>
        /// Type of the barchart
        /// </summary>
        public BarChartType BarChartType { get; set; }

        /// <summary>
        /// Direction of bar chart
        /// Horizontal = Bar chart (default)
        /// Vertical = Column chart
        /// </summary>
        public BarDirectionValues BarDirectionValues { get; set; } = BarDirectionValues.Bar;

        /// <summary>
        /// Type of bar grouping
        /// </summary>
        public BarGroupingValues BarGroupingValues { get; set; } = BarGroupingValues.Stacked;

        /// <summary>
        /// BarChart data source
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
        /// Categories
        /// </summary>
        public List<BarCategory> Categories { get; set; }

        /// <summary>
        /// Taille par défaut des textes dans le graphique
        /// Null par défaut
        /// </summary>
        public double? FontSize { get; set; } = null;

        /// <summary>
        /// Indicate if labels data will be showed
        /// </summary>
        public bool ShowDataLabel { get; set; } = true;

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
        /// Color of data labels
        /// </summary>
        public string DataLabelColor { get; set; }

        /// <summary>
        /// Values
        /// </summary>
        public List<BarSerie> Series { get; set; }

        /// <summary>
        /// Show / Hide Borders
        /// </summary>
        public bool ShowBarBorder { get; set; }

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
        /// Ctor
        /// </summary>
        public BarModel()
            : base(typeof(BarModel).Name)
        {
            BarChartType = BarChartType.BarChart;
        }
    }
}
