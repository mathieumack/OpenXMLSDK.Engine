using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.Bookmarks;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Bookmarks
{
    public class PlatformBookmarkEnd : PlatformOpenXmlElement, IBookmarkEnd
    {
        private readonly BookmarkEnd bookmarkEnd;

        public PlatformBookmarkEnd()
            : this(new BookmarkEnd())
        {
        }

        public PlatformBookmarkEnd(BookmarkEnd bookmarkEnd):
            base(bookmarkEnd)
        {
            this.bookmarkEnd = bookmarkEnd;
        }
    }
}
