namespace ReportEngine.Core.Template
{
    /// <summary>
    /// Start of the bookmark
    /// </summary>
    public class BookmarkStart : BaseElement
    {
        /// <summary>
        /// Id of the bookmark
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Name of the bookmark
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Bookmark end
        /// </summary>
        public BookmarkStart()
            : base(typeof(BookmarkStart).Name)
        {
        }
    }
}
