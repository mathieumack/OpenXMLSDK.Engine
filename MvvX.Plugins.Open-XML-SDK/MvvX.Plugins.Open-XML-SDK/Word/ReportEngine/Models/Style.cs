namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models
{
    /// <summary>
    /// Model class for a style in a word document
    /// </summary>
    public class Style : BaseElement
    {
        /// <summary>
        /// Id of style
        /// </summary>
        public string StyleId { get; set; }

        /// <summary>
        /// Type of style
        /// </summary>
        public StyleValues Type { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public Style()
            : base(typeof(Style).Name)
        {
        }
    }
}
