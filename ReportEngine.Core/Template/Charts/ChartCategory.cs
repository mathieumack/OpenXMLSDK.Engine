namespace ReportEngine.Core.Template.Charts
{
    public class ChartCategory : BaseElement
    {
        /// <summary>
        /// Name of the category
        /// </summary>
        public string Name { get; set; }

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
