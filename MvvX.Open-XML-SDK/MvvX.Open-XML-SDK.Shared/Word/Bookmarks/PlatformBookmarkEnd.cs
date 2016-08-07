using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word.Bookmarks;

namespace MvvX.Open_XML_SDK.Shared.Word.Bookmarks
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
