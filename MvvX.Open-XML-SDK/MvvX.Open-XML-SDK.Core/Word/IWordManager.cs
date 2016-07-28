using System;
using System.Collections.Generic;
using System.IO;
using MvvX.Open_XML_SDK.Core.Word.Bases;
using MvvX.Open_XML_SDK.Core.Word.Bookmarks;
using MvvX.Open_XML_SDK.Core.Word.Paragraphs;
using MvvX.Open_XML_SDK.Core.Word.Tables;
using MvvX.Open_XML_SDK.Core.Word.Tables.Models;

namespace MvvX.Open_XML_SDK.Core.Word
{
    public interface IWordManager : IDisposable
    {
        #region Bookmarks

        /// <summary>
        /// Permet de renvoyer la liste des signets du document Word
        /// </summary>
        /// <returns>Liste des noms des signets</returns>
        IList<string> GetBookmarks();

        /// <summary>
        /// Insère un texte au niveau des bookmark dont le nom correspond à szBookmarkPattern
        /// </summary>
        /// <param name="bookmark">Nom du signet</param>
        /// <param name="text"> Texte à insérer dans le signet (Bookmark)</param>
        /// <param name="textModel"> Propriété du texte : la couleur et le gras. Peut être à null</param>
        void SetTextOnBookmark(string bookmark, string text);

        /// <summary>
        /// Permet d'insérer un élément au niveau d'un bookmark
        /// </summary>
        /// <param name="bookmark">Nom du bookmark</param>
        /// <param name="element">Element à insérer</param>
        void SetOnBookmark(string bookmark, IOpenXmlElement element);

        /// <summary>
        /// Permet d'insérer un élément au niveau d'un bookmark
        /// </summary>
        /// <param name="bookmark">Nom du bookmark</param>
        /// <param name="paragraphs">Paragraphs à insérer</param>
        void SetParagraphsOnBookmark(string bookmark, IList<IParagraph> paragraphs);

        /// <summary>
        /// Permet de rechercher un bookmark dans le document :
        /// </summary>
        /// <param name="bookmark">Nom du bookmark</param>
        /// <returns></returns>
        IBookmarkEnd FindBookmark(string bookmark);

        #endregion

        #region Open / Save / Close

        /// <summary>
        /// Sauvegarde du document courant
        /// </summary>
        void SaveDoc();

        /// <summary>
        /// Fermeture du document avec sauvegarde automatique
        /// </summary>
        void CloseDocNoSave();

        /// <summary>
        /// Ouverture d'un document
        /// </summary>
        /// <param name="filePath">Chemin et nom complet du fichier à ouvrir</param>
        /// <param name="isEditable">Indique si le fichier doit être ouvert en mode éditable (Read/Write)</param>
        /// <returns>True si le document a bien été ouvert</returns>
        bool OpenDoc(string filePath, bool isEditable);

        /// <summary>
        /// Ouverture d'un document en mode streaming
        /// </summary>
        /// <param name="streamFile">Flux du fichier</param>
        /// <returns></returns>
        bool OpenDoc(Stream streamFile, bool isEditable);

        /// <summary>
        /// Ouverture d'un document depuis un template dotx
        /// </summary>
        /// <param name="streamTemplateFile">Chemin et nom complet du template</param>
        /// <param name="newFilePath">Chemin et nom complet du fichier qui sera sauvegardé</param>
        /// <returns>True si le document a bien été ouvert</returns>
        bool OpenDocFromTemplate(string templateFilePath);

        /// <summary>
        /// Ouverture d'un document depuis un template dotx
        /// </summary>
        /// <param name="templateFilePath">Chemin et nom complet du template</param>
        /// <param name="newFilePath">Chemin et nom complet du fichier qui sera sauvegardé</param>
        /// <param name="isEditable">Indique si le fichier doit être ouvert en mode éditable (Read/Write)</param>
        /// <returns>True si le document a bien été ouvert</returns>
        bool OpenDocFromTemplate(string templateFilePath, string newFilePath, bool isEditable);

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
