namespace ReportEngine.Core.Template.Charts
{
    public class BarCategory : BaseElement
    {
        /// <summary>
        /// Name of the category
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Category color
        /// </summary>
        public string Color { get; set; }

        public BarCategory() 
            : base(typeof(BarCategory).Name)
        {
        }
    }
}
