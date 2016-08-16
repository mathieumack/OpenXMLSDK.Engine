using System;
using System.Collections.Generic;
using System.IO;
using MvvX.Plugins.Open_XML_SDK.Core.Word.Bookmarks;
using MvvX.Plugins.Open_XML_SDK.Core.Word.Paragraphs;
using MvvX.Plugins.Open_XML_SDK.Core.Word.Tables;
using MvvX.Plugins.Open_XML_SDK.Core.Word.Tables.Models;
using MvvX.Plugins.Open_XML_SDK.Core.Word.Models;

namespace MvvX.Plugins.Open_XML_SDK.Core.Word
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
        /// Close the opened document
        /// </summary>
        void CloseDoc();

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

        #region Create IOpenXmlElement

        IRun CreateRun();

        ITable CreateTable();

        #endregion

        #region Texts
        
        IParagraph CreateParagraphForRun(IRun run);

        IRun CreateRunForTable(ITable run);

        IRun CreateRunForText(string content, RunPropertiesModel rpm = null);

        #endregion

        #region Tables

        /// <summary>
        /// Permet de créer une table
        /// </summary>
        /// <param name="properties"></param>
        /// <param name="headerCells"></param>
        /// <param name="cells"></param>
        /// <returns></returns>
        ITable CreateTable(IList<ITableRow> rows, TablePropertiesModel properties = null);

        /// <summary>
        /// Permet de créer un objet Cellule de tableau
        /// </summary>
        /// <param name="cellContent">Contenu de la cellule</param>
        /// <param name="width">Largeur de la cellule s'il y a lieu</param>
        /// <param name="shading">Informations de rendu de la cellule (couleur texte, fond ...)</param>
        /// <param name="justification">Justification à appliquer à la cellule</param>
        /// <param name="fusion">Indique si il faut fusionner la cellule</param>
        /// <param name="fusionChild">Indique si la cellule a fusionner et un enfant ou non de la première cellule fusionner</param>
        /// <returns></returns>
        ITableCell CreateTableCell(IRun cellContent, TableCellPropertiesModel cellModel);

        /// <summary>
        /// Permet de créer un objet Cellule de tableau
        /// </summary>
        /// <param name="cellContent">Contenus de la cellule</param>
        /// <param name="width">Largeur de la cellule s'il y a lieu</param>
        /// <param name="shading">Informations de rendu de la cellule (couleur texte, fond ...)</param>
        /// <param name="justification">Justification à appliquer à la cellule</param>
        /// <param name="gridSpan">Nombre de cellules horizontales fusionnées que représente la cellule</param>
        /// <returns></returns>
        ITableCell CreateTableCell(IList<IRun> cellContents, TableCellPropertiesModel cellModel);

        /// <summary>
        /// Permet de créer un objet Cellule de tableau
        /// </summary>
        /// <param name="cellContent">Contenus de la cellule</param>
        /// <param name="width">Largeur de la cellule s'il y a lieu</param>
        /// <param name="shading">Informations de rendu de la cellule (couleur texte, fond ...)</param>
        /// <param name="justification">Justification à appliquer à la cellule</param>
        /// <param name="gridSpan">Nombre de cellules horizontales fusionnées que représente la cellule</param>
        /// <returns></returns>
        ITableCell CreateTableCell(IList<IParagraph> cellContents, TableCellPropertiesModel cellModel);

        /// <summary>
        /// Permet de créer une cellule de tableau qui sera merger
        /// </summary>
        /// <param name="cellContents">Contenus de la cellule</param>
        /// <param name="width">Largeur de la cellule s'il y a lieu</param>
        /// <param name="shading">Informations de rendu de la cellule (couleur texte, fond ...)</param>
        /// <param name="Justification">Justification à appliquer à la cellule</param>
        /// <param name="fusion"></param>
        /// <param name="fusionChild"></param>
        /// <returns></returns>
        ITableCell CreateTableMergeCell(IRun cellContent, TableCellPropertiesModel cellModel);

        /// <summary>
        /// Permet de créer une cellule de tableau qui sera merger
        /// </summary>
        /// <param name="cellContents">Contenus de la cellule</param>
        /// <param name="width">Largeur de la cellule s'il y a lieu</param>
        /// <param name="shading">Informations de rendu de la cellule (couleur texte, fond ...)</param>
        /// <param name="Justification">Justification à appliquer à la cellule</param>
        /// <param name="fusion"></param>
        /// <param name="fusionChild"></param>
        /// <returns></returns>
        ITableCell CreateTableMergeCell(IList<IRun> cellContents, TableCellPropertiesModel cellModel);

        /// <summary>
        /// Permet de créer une ligne de tableau
        /// </summary>
        /// <param name="cells">Liste des cellules qui forment la ligne</param>
        /// <param name="tableRowProperties">Propriétées de la ligne</param>
        /// <returns></returns>
        ITableRow CreateTableRow(IList<ITableCell> cells, TableRowPropertiesModel properties = null);

        #endregion
    }
}
