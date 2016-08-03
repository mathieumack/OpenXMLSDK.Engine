using System;
using System.Collections.Generic;
using System.IO;
using MvvX.Open_XML_SDK.Core.Word.Bases;
using MvvX.Open_XML_SDK.Core.Word.Bookmarks;
using MvvX.Open_XML_SDK.Core.Word.Images;
using MvvX.Open_XML_SDK.Core.Word.Paragraphs;

namespace MvvX.Open_XML_SDK.Core.Word
{
    public interface IWordManager : IDisposable
    {
        #region Bookmarks

        /// <summary>
        /// Get bookmark list of doc
        /// </summary>
        IList<string> GetBookmarks();

        /// <summary>
        /// Insert text to bookmark
        /// </summary>
        /// <param name="bookmark">Bookmark name</param>
        /// <param name="text"> Text to insert</param>
        void SetTextOnBookmark(string bookmark, string text);

        /// <summary>
        /// Insert element in bookmark
        /// </summary>
        /// <param name="bookmark">Bookmark name</param>
        /// <param name="element">Element to insert</param>
        void SetOnBookmark(string bookmark, IOpenXmlElement element);

        /// <summary>
        /// Insert paragraph in bookmark
        /// </summary>
        /// <param name="bookmark">Bookmark name</param>
        /// <param name="paragraphs">Paragraphs to insert</param>
        void SetParagraphsOnBookmark(string bookmark, IList<IParagraph> paragraphs);

        /// <summary>
        /// Find bookmark in doc :
        /// </summary>
        /// <param name="bookmark">Bookmark name</param>
        /// <returns></returns>
        IBookmarkEnd FindBookmark(string bookmark);

        #endregion

        #region Open / Save / Close

        /// <summary>
        /// Save document
        /// </summary>
        void SaveDoc();

        /// <summary>
        /// Close document
        /// </summary>
        void CloseDocNoSave();

        /// <summary>
        /// Open doc
        /// </summary>
        /// <param name="filePath">Path and name of file to open</param>
        /// <param name="isEditable">Is file open in edit mode (Read/Write)</param>
        /// <returns>True if doc opened</returns>
        bool OpenDoc(string filePath, bool isEditable);

        /// <summary>
        /// Open doc in streaming mode
        /// </summary>
        /// <param name="streamFile">File flux</param>
        /// <returns></returns>
        bool OpenDoc(Stream streamFile, bool isEditable);

        /// <summary>
        /// Open document from template dotx
        /// </summary>
        /// <param name="streamTemplateFile">Template path and name</param>
        /// <param name="newFilePath">Path and name of generated file</param>
        /// <returns>True if doc opened</returns>
        bool OpenDocFromTemplate(string templateFilePath);

        /// <summary>
        /// Open document from template dotx
        /// </summary>
        /// <param name="templateFilePath">Template path</param>
        /// <param name="newFilePath">Path and name of generated file</param>
        /// <param name="isEditable">Is file open in edit mode (Read/Write)</param>
        /// <returns>True si le document a bien été ouvert</returns>
        bool OpenDocFromTemplate(string templateFilePath, string newFilePath, bool isEditable);

        #endregion

        #region Images

        /// <summary>
        /// Insert image to bookmark
        /// <param name="bookmark">Bookmark name</param>
        /// <param name="fileName">Image path and name</param>
        /// <param name="type">Image extension (.png, .jpg...)</param>
        /// </summary>
        void InsertPictureToBookmark(string bookmark, string fileName, ImageType type, long? maxWidth = null, long? maxHeight = null);
        
        #endregion
    }
}
