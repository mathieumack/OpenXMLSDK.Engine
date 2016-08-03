using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word.Bases;

namespace MvvX.Open_XML_SDK.Core.Word.Bookmarks
{
    public class PlatformBookmarkEnd : PlatformOpenXmlElement, IBookmarkEnd
    {
        private readonly BookmarkEnd bookmarkEnd;
        public PlatformBookmarkEnd(BookmarkEnd bookmarkEnd):
            base(bookmarkEnd)
        {
            this.bookmarkEnd = bookmarkEnd;
        }

        #region Static helpers methods

        public static PlatformBookmarkEnd New()
        {
            return new PlatformBookmarkEnd(new BookmarkEnd());
        }

        #endregion
    }
}
