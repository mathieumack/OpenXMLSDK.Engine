namespace ReportEngine.Core.Template
{
    public class HtmlContent : BaseElement
    {
        /// <summary>
        /// Html content (can contains #key# from context)
        /// </summary>
        public string Text { get; set; }

        public HtmlContent()
            :base(typeof(HtmlContent).Name)
        {

        }
    }
}
