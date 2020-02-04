namespace ReportEngine.Core.Template
{
    public class ForEachPage : Page
    {
        /// <summary>
        /// Define the prefix that will be used for automatically added items
        /// IsFirstItem, ...
        /// </summary>
        public string AutoContextAddItemsPrefix { get; set; }

        public string DataSourceKey { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public ForEachPage()
            : base(typeof(ForEachPage).Name)
        {
        }
    }
}
