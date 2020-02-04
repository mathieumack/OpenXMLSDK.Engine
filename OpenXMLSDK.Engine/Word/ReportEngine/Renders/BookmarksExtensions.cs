using DocumentFormat.OpenXml;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.Extensions;
using System;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    public static class BookmarksExtensions
    {
        /// <summary>
        /// Render a bookmarkStart element.
        /// </summary>
        /// <param name="hyperlink"></param>
        /// <param name="parent"></param>
        /// <param name="context"></param>
        /// <param name="documentPart"></param>
        /// <param name="formatProvider"></param>
        /// <returns></returns>
        public static OpenXmlElement Render(this BookmarkStart bookmarkStart, OpenXmlElement parent, ContextModel context, IFormatProvider formatProvider)
        {
            context.ReplaceItem(bookmarkStart, formatProvider);

            if (bookmarkStart.Show)
            {
                DocumentFormat.OpenXml.Wordprocessing.BookmarkStart bookmarkStartElement = new DocumentFormat.OpenXml.Wordprocessing.BookmarkStart()
                {
                    Id = bookmarkStart.Id,
                    Name = bookmarkStart.Name
                };

                parent.Append(bookmarkStartElement);

                return bookmarkStartElement;
            }

            return null;
        }

        /// <summary>
        /// Render a bookmarkStart element.
        /// </summary>
        /// <param name="hyperlink"></param>
        /// <param name="parent"></param>
        /// <param name="context"></param>
        /// <param name="documentPart"></param>
        /// <param name="formatProvider"></param>
        /// <returns></returns>
        public static OpenXmlElement Render(this BookmarkEnd bookmarkEnd, OpenXmlElement parent, ContextModel context, IFormatProvider formatProvider)
        {
            context.ReplaceItem(bookmarkEnd, formatProvider);

            if (bookmarkEnd.Show)
            {
                DocumentFormat.OpenXml.Wordprocessing.BookmarkEnd bookmarkEndElement = new DocumentFormat.OpenXml.Wordprocessing.BookmarkEnd()
                {
                    Id = bookmarkEnd.Id
                };

                parent.Append(bookmarkEndElement);

                return bookmarkEndElement;
            }

            return null;
        }
    }
}
