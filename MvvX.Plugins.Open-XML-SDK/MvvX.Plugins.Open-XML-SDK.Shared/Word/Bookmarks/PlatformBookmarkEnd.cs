using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.Open_XML_SDK.Core.Word.Bookmarks;

namespace MvvX.Plugins.Open_XML_SDK.Shared.Word.Bookmarks
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
