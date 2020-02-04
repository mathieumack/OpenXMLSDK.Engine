namespace ReportEngine.Core.Template
{
    /// <summary>
    /// Model class for header
    /// </summary>
    public class Header : BaseElement
    {
        /// <summary>
        /// Header type 
        /// </summary>
        public HeaderFooterValues Type { get; set; } = HeaderFooterValues.Default;

        /// <summary>
        /// Constructor
        /// </summary>
        public Header()
            : base(typeof(Header).Name)
        {
        }
    }
}
