namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models
{
    public class ForEachPage : Page
    {
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
