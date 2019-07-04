namespace OpenXMLSDK.Engine.Word.ReportEngine.Models
{
    /// <summary>
    /// End of the bookmark
    /// </summary>
    public class BookmarkEnd : BaseElement
    {
        /// <summary>
        /// Id of the bookmark
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Bookmark end
        /// </summary>
        public BookmarkEnd()
            : base(typeof(BookmarkEnd).Name)
        {
        }
    }
}
