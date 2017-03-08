namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models
{
    /// <summary>
    /// Model for a label
    /// </summary>
    public class Label : BaseElement
    {
        /// <summary>
        /// Label content (can contains #key# from context)
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        /// The label contains the page number
        /// </summary>
        public bool IsPageNumber { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public Label()
            : base(typeof(Label).Name)
        {
        }
    }
}
