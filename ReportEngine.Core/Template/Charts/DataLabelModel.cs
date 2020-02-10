namespace ReportEngine.Core.Template.Charts
{
    public class DataLabelModel
    {
        /// <summary>
        /// Indicate if labels data will be showed
        /// </summary>
        public bool ShowDataLabel { get; set; } = true;

        /// <summary>
        /// Indicate if category name will be showed
        /// </summary>
        public bool ShowCatName { get; set; } = true;

        /// <summary>
        /// Indicate if percentage will be showed
        /// </summary>
        public bool ShowPercent { get; set; } = true;

        /// <summary>
        /// Indicate position of label
        /// </summary>
        public DataLabelPositionValues LabelPosition { get; set; } = DataLabelPositionValues.BestFit;
        
        /// <summary>
        /// Indicate separator of differents values (label, data, percent...)
        /// </summary>
        public string Separator { get; set; } = string.Empty;

        /// <summary>
        /// Indicate font size of label 
        /// Instead of the string property, don't multiply by 2. This value is the real size ! 
        /// </summary>
        public int FontSize { get; set; } = 10;
    }
}
