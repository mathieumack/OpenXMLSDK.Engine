using System;
using System.IO;

namespace MvvX.Plugins.OpenXMLSDK.PowerPoint
{
    public interface IPowerPointManager : IDisposable
    {
        #region Open / Save / Close

        /// <summary>
        /// Close the opened document
        /// </summary>
        void CloseDoc();

        /// <summary>
        /// Get Stream for current document
        /// </summary>
        /// <returns></returns>
        MemoryStream GetMemoryStream();

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
    }
}
