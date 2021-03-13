namespace OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts
{
    public class ChartCategory : BaseElement
    {
        /// <summary>
        /// Name of the category
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Category value
        /// </summary>
        public double? Value { get; set; }

        /// <summary>
        /// Category color
        /// </summary>
        public string Color { get; set; }

        public ChartCategory()
            : base(typeof(ChartCategory).Name)
        {
        }
    }
}
