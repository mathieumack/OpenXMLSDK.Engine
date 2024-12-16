using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OpenXMLSDK.Engine.Word.Bookmarks
{
    public static class BookmarkExtensions
    {
        /// <summary>
        /// Finds the end of a bookmark by its name.
        /// </summary>
        /// <param name="wordManager">Word document manager</param>
        /// <param name="bookmark">Name of the bookmark</param>
        /// <returns>BookmarkEnd element if found, otherwise null</returns>
        public static BookmarkEnd FindBookmark(this WordManager wordManager, string bookmark)
        {
            var resultMain = wordManager.MainDocumentPart.RootElement.Descendants<BookmarkStart>().FirstOrDefault(e => e.Name == bookmark);
            if (resultMain != default(BookmarkStart))
                return wordManager.MainDocumentPart.RootElement.Descendants<BookmarkEnd>().FirstOrDefault(e => e.Id.Value == resultMain.Id);

            if (wordManager.MainDocumentPart.HeaderParts != null)
            {
                foreach (var header in wordManager.MainDocumentPart.HeaderParts)
                {
                    var result = header.RootElement.Descendants<BookmarkStart>().FirstOrDefault(e => e.Name == bookmark);
                    if (result != default(BookmarkStart))
                        return header.RootElement.Descendants<BookmarkEnd>().FirstOrDefault(e => e.Id.Value == result.Id);
                }
            }

            if (wordManager.MainDocumentPart.FooterParts != null)
            {
                foreach (var footer in wordManager.MainDocumentPart.FooterParts)
                {
                    var result = footer.RootElement.Descendants<BookmarkStart>().FirstOrDefault(e => e.Name == bookmark);
                    if (result != default(BookmarkStart))
                        return footer.RootElement.Descendants<BookmarkEnd>().FirstOrDefault(e => e.Id.Value == result.Id);
                }
            }

            return null;
        }

        /// <summary>
        /// Gets a list of all bookmark names in the document.
        /// </summary>
        /// <param name="wordManager">Word document manager</param>
        /// <returns>List of bookmark names</returns>
        public static IList<string> GetBookmarks(this WordManager wordManager)
        {
            if (wordManager.WordprocessingDocument == null)
                throw new InvalidOperationException("Document not loaded");

            var result = wordManager.MainDocumentPart.Document.Body.Descendants<BookmarkStart>().Select(e => e.Name.Value).ToList();
            if (wordManager.MainDocumentPart.HeaderParts != null)
            {
                foreach (var header in wordManager.MainDocumentPart.HeaderParts)
                    result.AddRange(header.RootElement.Descendants<BookmarkStart>().Select(e => e.Name.Value).ToList());
            }

            if (wordManager.MainDocumentPart.FooterParts != null)
            {
                foreach (var footer in wordManager.MainDocumentPart.FooterParts)
                    result.AddRange(footer.RootElement.Descendants<BookmarkStart>().Select(e => e.Name.Value).ToList());
            }
            return result;
        }

        /// <summary>
        /// Inserts an OpenXmlElement at the specified bookmark.
        /// </summary>
        /// <param name="wordManager">Word document manager</param>
        /// <param name="bookmark">Name of the bookmark</param>
        /// <param name="element">Element to insert</param>
        public static void SetOnBookmark(this WordManager wordManager, string bookmark, OpenXmlElement element)
        {
            if (string.IsNullOrWhiteSpace(bookmark))
                throw new ArgumentNullException(nameof(bookmark), "Bookmark must not be null or white spaces");
            if (wordManager.WordprocessingDocument == null)
                throw new InvalidOperationException("Document not loaded");

            var bookmarkElement = wordManager.FindBookmark(bookmark);
            if (bookmarkElement != null)
                bookmarkElement.InsertAfterSelf(element);
        }

        /// <summary>
        /// Inserts a list of paragraphs at the specified bookmark.
        /// </summary>
        /// <param name="wordManager">Word document manager</param>
        /// <param name="bookmark">Name of the bookmark</param>
        /// <param name="paragraphs">List of paragraphs to insert</param>
        public static void SetParagraphsOnBookmark(this WordManager wordManager, string bookmark, IList<Paragraph> paragraphs)
        {
            if (string.IsNullOrWhiteSpace(bookmark))
                throw new ArgumentNullException(nameof(bookmark), "Bookmark must not be null or white spaces");
            if (wordManager.WordprocessingDocument == null)
                throw new InvalidOperationException("Document not loaded");

            var bookmarkElement = wordManager.FindBookmark(bookmark);
            if (bookmarkElement != null)
            {
                var paragraph = bookmarkElement.Ancestors<Paragraph>().LastOrDefault();
                var firstParagraph = bookmarkElement.Ancestors<Paragraph>().LastOrDefault();
                if (paragraph != null)
                {
                    foreach (var item in paragraphs)
                        paragraph = paragraph.InsertAfterSelf(item);

                    if (paragraphs.Any())
                        firstParagraph.Remove();
                }
            }
        }

        /// <summary>
        /// Inserts text at the specified bookmark.
        /// </summary>
        /// <param name="wordManager">Word document manager</param>
        /// <param name="bookmark">Name of the bookmark</param>
        /// <param name="text">Text to insert</param>
        public static void SetTextOnBookmark(this WordManager wordManager, string bookmark, string text)
        {
            if (string.IsNullOrWhiteSpace(bookmark))
                throw new ArgumentNullException(nameof(bookmark), "Bookmark must not be null or white spaces");
            if (wordManager.WordprocessingDocument == null)
                throw new InvalidOperationException("Document not loaded");

            var run = new Run(new Text(text));
            wordManager.SetOnBookmark(bookmark, run);
        }

        /// <summary>
        /// Inserts a list of texts at the specified bookmark.
        /// </summary>
        /// <param name="wordManager">Word document manager</param>
        /// <param name="bookmark">Name of the bookmark</param>
        /// <param name="texts">List of texts to insert</param>
        /// <param name="formated">Indicates if the texts should be formatted with line breaks</param>
        public static void SetTextsOnBookmark(this WordManager wordManager, string bookmark, List<string> texts, bool formated = true)
        {
            if (string.IsNullOrWhiteSpace(bookmark))
                throw new ArgumentNullException(nameof(bookmark), "Bookmark must not be null or white spaces");
            if (wordManager.WordprocessingDocument == null)
                throw new InvalidOperationException("Document not loaded");

            var run = new Run();

            for (int i = 0; i < texts.Count; i++)
            {
                run.Append(new Text() { Text = texts[i] });

                if (i < texts.Count - 1 && formated)
                    run.Append(new Break());
                else if (!formated)
                    run.Append(new Text() { Text = " ", Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve });
            }

            wordManager.SetOnBookmark(bookmark, run);
        }

        /// <summary>
        /// Inserts HTML content at the specified bookmark.
        /// </summary>
        /// <param name="wordManager">Word document manager</param>
        /// <param name="bookmark">Name of the bookmark</param>
        /// <param name="html">HTML content to insert</param>
        public static void SetHtmlOnBookmark(this WordManager wordManager, string bookmark, string html)
        {
            if (string.IsNullOrWhiteSpace(bookmark))
                throw new ArgumentNullException(nameof(bookmark), "Bookmark must not be null or white spaces");
            if (wordManager.WordprocessingDocument == null)
                throw new InvalidOperationException("Document not loaded");

            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(html)))
            {
                wordManager.AddAltChunkOnBookmark(bookmark, ms, AlternativeFormatImportPartType.Xhtml);
            }
        }

        /// <summary>
        /// Inserts another Word document at the specified bookmark.
        /// </summary>
        /// <param name="wordManager">Word document manager</param>
        /// <param name="bookmark">Name of the bookmark</param>
        /// <param name="content">Stream of the Word document to insert</param>
        public static void SetSubDocumentOnBookmark(this WordManager wordManager, string bookmark, Stream content)
        {
            if (string.IsNullOrWhiteSpace(bookmark))
                throw new ArgumentNullException(nameof(bookmark), "Bookmark must not be null or white spaces");
            if (wordManager.WordprocessingDocument == null)
                throw new InvalidOperationException("Document not loaded");

            wordManager.AddAltChunkOnBookmark(bookmark, content, AlternativeFormatImportPartType.WordprocessingML);
        }

        /// <summary>
        /// Inserts a document at the specified bookmark.
        /// </summary>
        /// <param name="wordManager">Word document manager</param>
        /// <param name="bookmark">Name of the bookmark</param>
        /// <param name="content">Stream of the document to insert</param>
        /// <param name="importType">Type of the document to insert</param>
        private static void AddAltChunkOnBookmark(this WordManager wordManager, string bookmark, Stream content, PartTypeInfo importType)
        {
            var formatImportPart = wordManager.MainDocumentPart.AddAlternativeFormatImportPart(importType);

            formatImportPart.FeedData(content);

            AltChunk altChunk = new AltChunk();
            altChunk.Id = wordManager.MainDocumentPart.GetIdOfPart(formatImportPart);

            var bookmarkElement = wordManager.FindBookmark(bookmark);
            if (bookmarkElement != null)
            {
                var paragraph = bookmarkElement.Ancestors<Paragraph>().LastOrDefault();
                // without an empty paragraph after altchunk, the docx might be corrupted if the bookmark is inside a table and the html only contains one paragraph
                if (paragraph.Ancestors<Table>().Any())
                    paragraph.InsertAfterSelf(new Paragraph());
                paragraph.InsertAfterSelf(new AltChunk(altChunk));
            }
        }

        /// <summary>
        /// Merges documents and inserts them at the specified bookmark.
        /// </summary>
        /// <param name="wordManager">Word document manager</param>
        /// <param name="bookmark">Name of the bookmark</param>
        /// <param name="filesToInsert">List of documents to insert</param>
        /// <param name="insertPageBreaks">Indicates if a page break should be added after each document</param>
        public static void InsertDocsToBookmark(this WordManager wordManager, string bookmark, IList<MemoryStream> filesToInsert, bool insertPageBreaks)
        {
            if (string.IsNullOrWhiteSpace(bookmark))
                throw new ArgumentNullException(nameof(bookmark), "Bookmark must not be null or white spaces");
            if (filesToInsert == null)
                throw new ArgumentNullException(nameof(filesToInsert), "FilesToInsert must not be null");

            var bookmarkElement = wordManager.FindBookmark(bookmark);
            if (bookmarkElement != default(BookmarkEnd))
            {
                OpenXmlCompositeElement insertAfterElement = wordManager.WordprocessingDocument.MainDocumentPart.Document.Body.Descendants<BookmarkStart>().SingleOrDefault<BookmarkStart>((BookmarkStart b) => b.Name == bookmark)
                    .Ancestors<Paragraph>().FirstOrDefault<Paragraph>();
                wordManager.AppendStreams(filesToInsert, insertPageBreaks, insertAfterElement);
            }
        }
    }
}
