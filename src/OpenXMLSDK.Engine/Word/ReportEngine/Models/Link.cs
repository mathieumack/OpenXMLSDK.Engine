namespace OpenXMLSDK.Engine.Word.ReportEngine.Models
{
    /// <summary>
    /// Link is a sub-model, allowing the configuration of link inside some BaseElement
    /// </summary>
    public class Link
    {
        /// <summary>
        /// Define a redirection link when clicking on the picture
        /// </summary>
        public string HyperlinkUrl { get; set; }

        /// <summary>
        /// Define if URL is external
        /// </summary>
        public bool IsExternalUrl { get; set; }
    }
}
