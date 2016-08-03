using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word;
using MvvX.Open_XML_SDK.Core.Word.Bases;
using MvvX.Open_XML_SDK.Core.Word.Bookmarks;
using MvvX.Open_XML_SDK.Core.Word.Paragraphs;
using MvvX.Open_XML_SDK.Core.Word.Tables;
using MvvX.Open_XML_SDK.Core.Word.Tables.Models;
using MvvX.Open_XML_SDK.Shared.Word.Extensions;
using DocumentFormat.OpenXml;
using MvvX.Open_XML_SDK.Shared.Word;

namespace MvvX.Open_XML_SDK.Word
{
    public class WordManager : IWordManager
    {
        #region Fields

        private MemoryStream streamFile;

        private string filePath;

        /// <summary>
        /// Indique si le document en cours l'est en mode d'édition
        /// </summary>
        private bool isEditable = false;

        /// <summary>
        /// Objet Word courant
        /// </summary>
        private WordprocessingDocument wdDoc = null;

        /// <summary>
        /// Page principale
        /// </summary>
        private MainDocumentPart wdMainDocumentPart = null;

        #endregion

        #region Constructeurs

        /// <summary>
        /// 
        /// </summary>
        public WordManager()
        {
            filePath = string.Empty;
        }

        #endregion

        #region Dispose / fin

        public void CloseDoc()
        {
            this.SaveDoc();
            wdDoc.Close();
        }

        public void Dispose()
        {
            if (wdDoc != null)
                wdDoc.Dispose();
        }

        /// <summary>
        /// Permet de renvoyer le MemoryStream associé au document en cours
        /// </summary>
        /// <returns>MemoryStream en cours, null sinon</returns>
        public MemoryStream GetMemoryStream()
        {
            var memoryStream = new MemoryStream();
            streamFile.Position = 0;
            streamFile.CopyTo(memoryStream);
            return memoryStream;
        }

        #endregion

        #region Fonctions privées - Open/Save/Create

        /// <summary>
        /// Sauvegarde du document courant
        /// </summary>
        public void SaveDoc()
        {
            if (wdDoc == null)
                throw new InvalidOperationException("Document not loaded");

            wdMainDocumentPart.Document.Save();
        }

        /// <summary>
        /// Fermeture du document avec sauvegarde automatique
        /// </summary>
        public void CloseDocNoSave()
        {
            if (wdDoc == null)
                throw new InvalidOperationException("Document not loaded");

            wdDoc.Close();
        }

        /// <summary>
        /// Ouverture d'un document
        /// </summary>
        /// <param name="filePath">Chemin et nom complet du fichier à ouvrir</param>
        /// <param name="isEditable">Indique si le fichier doit être ouvert en mode éditable (Read/Write)</param>
        /// <returns>True si le document a bien été ouvert</returns>
        public bool OpenDoc(string filePath, bool isEditable)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentNullException("filePath must be not null or white spaces");
            if (!File.Exists(filePath))
                throw new FileNotFoundException("file not found");

            this.isEditable = isEditable;

            try
            {
                wdDoc = WordprocessingDocument.Open(filePath, isEditable);
                wdMainDocumentPart = wdDoc.MainDocumentPart;

                return true;
            }
            catch (Exception ex)
            {
                wdDoc = null;
                return false;
            }
        }

        /// <summary>
        /// Ouverture d'un document en mode streaming
        /// </summary>
        /// <param name="streamFile">Flux du fichier</param>
        /// <returns></returns>
        public bool OpenDoc(Stream streamFile, bool isEditable)
        {
            if (streamFile == null)
                throw new ArgumentNullException("streamFile must be not null");

            this.isEditable = isEditable;

            try
            {
                wdDoc = WordprocessingDocument.Open(streamFile, isEditable);
                wdMainDocumentPart = wdDoc.MainDocumentPart;

                return true;
            }
            catch (Exception ex)
            {
                wdDoc = null;
                return false;
            }
        }

        /// <summary>
        /// Ouverture d'un document depuis un template dotx
        /// </summary>
        /// <param name="streamTemplateFile">Chemin et nom complet du template</param>
        /// <param name="newFilePath">Chemin et nom complet du fichier qui sera sauvegardé</param>
        /// <returns>True si le document a bien été ouvert</returns>
        public bool OpenDocFromTemplate(string templateFilePath)
        {
            if (string.IsNullOrWhiteSpace(templateFilePath))
                throw new ArgumentNullException("templateFilePath must be not null or white spaces");
            if (!File.Exists(templateFilePath))
                throw new FileNotFoundException("file not found");

            this.isEditable = true;

            streamFile = new MemoryStream();
            try
            {
                byte[] byteArray = File.ReadAllBytes(templateFilePath);
                streamFile.Write(byteArray, 0, (int)byteArray.Length);

                wdDoc = WordprocessingDocument.Open(streamFile, isEditable);

                // Change the document type to Document
                wdDoc.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

                wdMainDocumentPart = wdDoc.MainDocumentPart;

                return true;
            }
            catch (Exception ex)
            {
                wdDoc = null;
                return false;
            }
        }

        /// <summary>
        /// Ouverture d'un document depuis un template dotx
        /// </summary>
        /// <param name="templateFilePath">Chemin et nom complet du template</param>
        /// <param name="newFilePath">Chemin et nom complet du fichier qui sera sauvegardé</param>
        /// <param name="isEditable">Indique si le fichier doit être ouvert en mode éditable (Read/Write)</param>
        /// <returns>True si le document a bien été ouvert</returns>
        public bool OpenDocFromTemplate(string templateFilePath, string newFilePath, bool isEditable)
        {
            if (string.IsNullOrWhiteSpace(templateFilePath))
                throw new ArgumentNullException("templateFilePath must be not null or white spaces");
            if (!File.Exists(templateFilePath))
                throw new FileNotFoundException("file not found");

            this.isEditable = isEditable;

            filePath = newFilePath;
            try
            {
                System.IO.File.Copy(templateFilePath, newFilePath, true);

                wdDoc = WordprocessingDocument.Open(filePath, isEditable);

                // Change the document type to Document
                wdDoc.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

                wdMainDocumentPart = wdDoc.MainDocumentPart;

                SaveDoc();
                CloseDocNoSave();

                return OpenDoc(newFilePath, isEditable);
            }
            catch (Exception ex)
            {
                wdDoc = null;
                return false;
            }
        }

        #endregion

        #region Bookmarks

        public IBookmarkEnd FindBookmark(string bookmark)
        {
            var resultMain = wdMainDocumentPart.RootElement.Descendants<BookmarkStart>().FirstOrDefault(e => e.Name == bookmark);
            if (resultMain != default(BookmarkStart))
                return new PlatformBookmarkEnd(wdMainDocumentPart.RootElement.Descendants<BookmarkEnd>().FirstOrDefault(e => e.Id.Value == resultMain.Id));

            if (wdMainDocumentPart.HeaderParts != null)
            {
                foreach (var header in wdMainDocumentPart.HeaderParts)
                {
                    var result = header.RootElement.Descendants<BookmarkStart>().FirstOrDefault(e => e.Name == bookmark);
                    var resultBE = header.RootElement.Descendants<BookmarkEnd>().Select(e => e.Id).ToList();
                    if (result != default(BookmarkStart))
                        return new PlatformBookmarkEnd(header.RootElement.Descendants<BookmarkEnd>().FirstOrDefault(e => e.Id.Value == result.Id));
                }
            }

            if (wdMainDocumentPart.FooterParts != null)
            {
                foreach (var footer in wdMainDocumentPart.FooterParts)
                {
                    var result = footer.RootElement.Descendants<BookmarkStart>().FirstOrDefault(e => e.Name == bookmark);
                    if (result != default(BookmarkStart))
                        return new PlatformBookmarkEnd(footer.RootElement.Descendants<BookmarkEnd>().FirstOrDefault(e => e.Id.Value == result.Id));
                }
            }

            return null;
        }

        public IList<string> GetBookmarks()
        {
            if (wdDoc == null)
                throw new InvalidOperationException("Document not loaded");

            var result = wdMainDocumentPart.Document.Body.Descendants<BookmarkStart>().Select(e => e.Name.Value).ToList();
            if (wdMainDocumentPart.HeaderParts != null)
            {
                foreach (var header in wdMainDocumentPart.HeaderParts)
                    result.AddRange(header.RootElement.Descendants<BookmarkStart>().Select(e => e.Name.Value).ToList());
            }

            if (wdMainDocumentPart.FooterParts != null)
            {
                foreach (var footer in wdMainDocumentPart.FooterParts)
                    result.AddRange(footer.RootElement.Descendants<BookmarkStart>().Select(e => e.Name.Value).ToList());
            }
            return result;
        }

        public void SetOnBookmark(string bookmark, IOpenXmlElement element)
        {
            if (string.IsNullOrWhiteSpace(bookmark))
                throw new ArgumentNullException("bookmark must be not null or white spaces");
            if (wdDoc == null)
                throw new InvalidOperationException("Document not loaded");

            var bookmarkElement = FindBookmark(bookmark);
            if (bookmarkElement != null)
                bookmarkElement.InsertAfterSelf(element);
        }

        public void SetParagraphsOnBookmark(string bookmark, IList<IParagraph> paragraphs)
        {
            if (string.IsNullOrWhiteSpace(bookmark))
                throw new ArgumentNullException("bookmark must be not null or white spaces");
            if (wdDoc == null)
                throw new InvalidOperationException("Document not loaded");

            var bookmarkElement = FindBookmark(bookmark);
            if (bookmarkElement != null)
            {
                var paragraph = bookmarkElement.Ancestors<IParagraph>().LastOrDefault();
                if (paragraph != null)
                {
                    foreach (var item in paragraphs)
                        paragraph = paragraph.InsertAfterSelf(item);
                }
            }
        }

        public void SetTextOnBookmark(string bookmark, string text)
        {
            if (string.IsNullOrWhiteSpace(bookmark))
                throw new ArgumentNullException("bookmark must be not null or white spaces");
            if (wdDoc == null)
                throw new InvalidOperationException("Document not loaded");

            var run = new Run(new Text(text));
            SetOnBookmark(bookmark, new PlatformRun(run));
        }

        #endregion

        #region Shadings

        public IShading GetShading(string textColor = null, string fillColor = null)
        {
            var shading = new Shading();

            if (!string.IsNullOrWhiteSpace(textColor))
                shading.Color = textColor;

            if (!string.IsNullOrWhiteSpace(fillColor))
                shading.Fill = fillColor;

            return new PlatformShading(shading);
        }

        #endregion

        #region Texts

        public IRun CreateEmptyRun()
        {
            return new PlatformRun(new Run());
        }

        public IRun CreateRunForTable(ITable table)
        {
            if (table == null)
                throw new ArgumentNullException("table must not be null");

            var run = new Run(table.ContentItem as Table);

            return new PlatformRun(run);
        }

        public IParagraph CreateParagraphForRun(IRun run)
        {
            if (run == null)
                throw new ArgumentNullException("run must not be null");

            var paragraph = new Paragraph(run.ContentItem as Run);

            return new PlatformParagraph(paragraph);
        }

        public IRun CreateRunForTexte(string content, RunPropertiesModel rpm = null)
        {
            RunProperties runHeader = new RunProperties();

            Run run = new Run();

            if (rpm != null)
            {
                if (rpm.Bold.HasValue)
                    runHeader.Append(new Bold() { Val = rpm.Bold.Value });
                if (rpm.Italic.HasValue)
                    runHeader.Append(new Italic() { Val = rpm.Italic.Value });
                if (!string.IsNullOrWhiteSpace(rpm.FontFamily))
                    runHeader.Append(new RunFonts()
                    {
                        Ascii = rpm.FontFamily,
                        HighAnsi = rpm.FontFamily,
                        EastAsia = rpm.FontFamily,
                        ComplexScript = rpm.FontFamily
                    });
                if (!string.IsNullOrWhiteSpace(rpm.FontSize))
                    runHeader.Append(new FontSize() { Val = rpm.FontSize });
                if (!string.IsNullOrWhiteSpace(rpm.Color))
                    runHeader.Append(new Color() { Val = rpm.Color });

                if (rpm.Bold.HasValue || rpm.Italic.HasValue || rpm.FontFamily != null || rpm.FontSize != null || rpm.Color != null)
                    run.Append(runHeader);
            }

            run.Append(new Text(content)
            {
                Space = SpaceProcessingModeValues.Preserve
            });

            return new PlatformRun(run);
        }

        #endregion

        #region Tables

        public ITable CreateTable(IList<ITableRow> rows, TablePropertiesModel properties)
        {
            if (rows == null)
                throw new ArgumentNullException("rows must not be null");
            if (rows.Any(e => e == null))
                throw new ArgumentNullException("All elements of rows must not be null");
            if (properties == null || properties.BottomBorder == null)
                throw new ArgumentNullException("properties and properties.BottomBorder must not be null");
            if (properties == null || properties.InsideHorizontalBorder == null)
                throw new ArgumentNullException("properties and properties.InsideHorizontalBorder must not be null");
            if (properties == null || properties.InsideVerticalBorder == null)
                throw new ArgumentNullException("properties and properties.InsideVerticalBorder must not be null");
            if (properties == null || properties.LeftBorder == null)
                throw new ArgumentNullException("properties and properties.LeftBorder must not be null");
            if (properties == null || properties.RightBorder == null)
                throw new ArgumentNullException("properties and properties.RightBorder must not be null");
            if (properties == null || properties.TopBorder == null)
                throw new ArgumentNullException("properties and properties.TopBorder must not be null");

            // Vérification que chaque ligne a bien le même nombre de colonne
            if (rows.Count > 1)
            {
                var nbCellRow1 = 0;
                var nbCellOtherRow = 0;
                if (rows[0].Descendants<ITableCell>().FirstOrDefault() != null)
                {
                    if (rows[0].Descendants<IGridSpan>().FirstOrDefault() != null)
                        nbCellRow1 = rows[0].Descendants<ITableCell>().Count() - rows[0].Descendants<IGridSpan>().Count() + rows[0].Descendants<IGridSpan>().Sum(e => e.Val);
                    // nb de Cell - nb de GridSpan + Somme des GridSpan
                    else
                        nbCellRow1 = rows[0].Descendants<ITableCell>().Count();
                    // nb de Cell 
                }
                for (int i = 1; i < rows.Count; i++)
                {
                    if (rows[i].Descendants<ITableCell>().FirstOrDefault() != default(TableCell))
                    {
                        if (rows[i].Descendants<IGridSpan>().FirstOrDefault() != default(GridSpan))
                            nbCellOtherRow = rows[i].Descendants<ITableCell>().Count() - rows[i].Descendants<IGridSpan>().Count() + rows[i].Descendants<IGridSpan>().Sum(e => e.Val);
                        // nb de Cell - nb de GridSpan + Somme des GridSpan
                        else
                            nbCellOtherRow = rows[i].Descendants<ITableCell>().Count();
                        // nb de Cell 
                    }

                    if (!nbCellRow1.Equals(nbCellOtherRow))
                        throw new ArgumentException("Lines 1 and " + (i + 1) + " have not same columns count");
                }
            }

            Table tbl = new Table();

            TableProperties tableProp = new TableProperties();
            TableStyle tableStyle = new TableStyle() { Val = "TableGrid" };
            // 100% de la largeur de la page ou largeur indiquée en paramètre:
            TableWidth tableWidth = new TableWidth();
            if (properties != null && !string.IsNullOrWhiteSpace(properties.Width))
            {
                tableWidth.Width = properties.Width;
                if (properties.WidthUnit.OOxmlEquals(DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Nil))
                    tableWidth.Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Pct;
                else
                    tableWidth.Type = properties.WidthUnit.ToOOxml();
            }
            else
            {
                tableWidth.Width = "5000";
                tableWidth.Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Pct;
            }
            tableProp.Append(tableStyle, tableWidth);
            tbl.AppendChild(tableProp);

            if (properties != null)
            {
                // Create a TableProperties object and specify its border information.
                TableProperties tblProp = new TableProperties(
                    new TableBorders(
                        new TopBorder() { Val = new EnumValue<DocumentFormat.OpenXml.Wordprocessing.BorderValues>(properties.TopBorder.BorderValue.ToOOxml()), Size = Convert.ToUInt32(properties.TopBorder.Size), Color = properties.TopBorder.Color },
                        new BottomBorder() { Val = new EnumValue<DocumentFormat.OpenXml.Wordprocessing.BorderValues>(properties.BottomBorder.BorderValue.ToOOxml()), Size = Convert.ToUInt32(properties.BottomBorder.Size), Color = properties.BottomBorder.Color },
                        new LeftBorder() { Val = new EnumValue<DocumentFormat.OpenXml.Wordprocessing.BorderValues>(properties.LeftBorder.BorderValue.ToOOxml()), Size = Convert.ToUInt32(properties.LeftBorder.Size), Color = properties.LeftBorder.Color },
                        new RightBorder() { Val = new EnumValue<DocumentFormat.OpenXml.Wordprocessing.BorderValues>(properties.RightBorder.BorderValue.ToOOxml()), Size = Convert.ToUInt32(properties.RightBorder.Size), Color = properties.RightBorder.Color },
                        new InsideHorizontalBorder() { Val = new EnumValue<DocumentFormat.OpenXml.Wordprocessing.BorderValues>(properties.InsideHorizontalBorder.BorderValue.ToOOxml()), Size = Convert.ToUInt32(properties.InsideHorizontalBorder.Size), Color = properties.InsideHorizontalBorder.Color },
                        new InsideVerticalBorder() { Val = new EnumValue<DocumentFormat.OpenXml.Wordprocessing.BorderValues>(properties.InsideVerticalBorder.BorderValue.ToOOxml()), Size = Convert.ToUInt32(properties.InsideVerticalBorder.Size), Color = properties.InsideVerticalBorder.Color }
                    ),
                    new TableProperties(
                        new TableJustification() { Val = properties.TableJustification.ToOOxml() }
                    )
                );
                // Append the TableProperties object to the empty table.
                tbl.AppendChild(tblProp);
            }

            if (rows != null)
            {
                for (int i = 0; i < rows.Count; i++)
                {
                    tbl.Append(rows[i].ContentItem as TableRow);
                }
            }

            return new PlatformTable(tbl);
        }

        public ITableCell CreateTableCell(IRun cellContent, TableCellPropertiesModel cellModel)
        {
            if (cellContent == null)
                throw new ArgumentNullException("cellContent must not be null");
            if (cellModel == null)
                throw new ArgumentNullException("cellModel must not be null");

            return CreateTableCell(new List<IRun>() { cellContent }, cellModel);
        }

        public ITableCell CreateTableCell(IList<IRun> cellContents, TableCellPropertiesModel cellModel)
        {
            if (cellContents == null)
                throw new ArgumentNullException("cellContents must not be null");
            if (cellContents.Count == 0)
                throw new ArgumentNullException("cellContents must not be an empty list");
            if (cellContents.Any(e => e == null))
                throw new ArgumentNullException("All elements of cellContents must be not null");
            if (cellModel == null)
                throw new ArgumentNullException("cellModel must not be null");

            TableCell tc = new TableCell();

            var tableCellProperties = new TableCellProperties();
            if (!string.IsNullOrWhiteSpace(cellModel.Width))
                if (cellModel.WidthUnit.Equals(DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Nil))
                    tableCellProperties.Append(new TableCellWidth { Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Dxa, Width = cellModel.Width });
                else
                    tableCellProperties.Append(new TableCellWidth { Type = cellModel.WidthUnit.ToOOxml(), Width = cellModel.Width });
            else
                tableCellProperties.Append(new TableCellWidth { Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Auto });

            if (cellModel.Gridspan.HasValue)
                tableCellProperties.Append(new GridSpan { Val = cellModel.Gridspan.Value });

            if (cellModel.Shading != null)
                tableCellProperties.Append(cellModel.Shading.ContentItem as Shading);

            tableCellProperties.Append(new TableCellVerticalAlignment { Val = cellModel.TableVerticalAlignementValues.ToOOxml() });

            // Modification de la rotation du texte dans la cellule
            if (cellModel.TextDirectionValues.HasValue)
                tableCellProperties.Append(new TextDirection { Val = cellModel.TextDirectionValues.ToOOxml() });

            // Gestion des bordures de la cellule
            if (cellModel.TopBorder != null)
                tableCellProperties.Append(new TableCellBorders(new TopBorder { Val = cellModel.TopBorder.BorderValue.ToOOxml(), Color = cellModel.TopBorder.Color, Size =Convert.ToUInt32(cellModel.TopBorder.Size) }));

            if (cellModel.BottomBorder != null)
                tableCellProperties.Append(new TableCellBorders(new BottomBorder { Val = cellModel.BottomBorder.BorderValue.ToOOxml(), Color = cellModel.BottomBorder.Color, Size = Convert.ToUInt32(cellModel.BottomBorder.Size) }));

            if (cellModel.LeftBorder != null)
                tableCellProperties.Append(new TableCellBorders(new LeftBorder { Val = cellModel.LeftBorder.BorderValue.ToOOxml(), Color = cellModel.LeftBorder.Color, Size = Convert.ToUInt32(cellModel.LeftBorder.Size) }));

            if (cellModel.RightBorder != null)
                tableCellProperties.Append(new TableCellBorders(new RightBorder { Val = cellModel.RightBorder.BorderValue.ToOOxml(), Color = cellModel.RightBorder.Color, Size = Convert.ToUInt32(cellModel.RightBorder.Size) }));

            tc.Append(tableCellProperties);

            Paragraph par = new Paragraph();
            ParagraphProperties pr = new ParagraphProperties(new TableCellVerticalAlignment { Val = cellModel.TableVerticalAlignementValues.ToOOxml() });
            //new SpacingBetweenLines() { After = cellModel.SpacingAfter, Before = cellModel.SpacingBefore, Line = "240" });

            if (cellModel.Justification.HasValue)
                pr.Append(new Justification() { Val = cellModel.Justification.Value.ToOOxml() });

            if (cellModel.ParagraphSolidarity)
                pr.Append(new KeepNext());
            par.Append(pr);

            for (int i = 0; i < cellContents.Count; i++)
                par.Append(cellContents[i].ContentItem as Run);

            tc.Append(par);

            return new PlatformTableCell(tc);
        }

        public ITableCell CreateTableCell(IList<IParagraph> cellContents, TableCellPropertiesModel cellModel)
        {
            if (cellContents == null)
                throw new ArgumentNullException("cellContents must not be null");
            if (cellContents.Count == 0)
                throw new ArgumentNullException("cellContents must not be an empty list");
            if (cellContents.Any(e => e == null))
                throw new ArgumentNullException("All elements of cellContents must be not null");
            if (cellModel == null)
                throw new ArgumentNullException("cellModel must not be null");

            TableCell tc = new TableCell();

            var tableCellProperties = new TableCellProperties();
            if (!string.IsNullOrWhiteSpace(cellModel.Width))
                if (cellModel.WidthUnit.OOxmlEquals(DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Nil))
                    tableCellProperties.Append(new TableCellWidth { Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Dxa, Width = cellModel.Width });
                else
                    tableCellProperties.Append(new TableCellWidth { Type = cellModel.WidthUnit.ToOOxml(), Width = cellModel.Width });
            else
                tableCellProperties.Append(new TableCellWidth { Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Auto });

            if (cellModel.Gridspan.HasValue)
                tableCellProperties.Append(new GridSpan { Val = cellModel.Gridspan.Value });

            if (cellModel.Shading != null)
                tableCellProperties.Append(cellModel.Shading.ContentItem as Shading);

            tableCellProperties.Append(new TableCellVerticalAlignment { Val = cellModel.TableVerticalAlignementValues.ToOOxml() });

            // Modification de la rotation du texte dans la cellule
            if (cellModel.TextDirectionValues.HasValue)
                tableCellProperties.Append(new TextDirection { Val = cellModel.TextDirectionValues.ToOOxml() });

            // Gestion des bordures de la cellule
            if (cellModel.TopBorder != null)
                tableCellProperties.Append(new TableCellBorders(new TopBorder { Val = cellModel.TopBorder.BorderValue.ToOOxml(), Color = cellModel.TopBorder.Color, Size = Convert.ToUInt32(cellModel.TopBorder.Size) }));

            if (cellModel.BottomBorder != null)
                tableCellProperties.Append(new TableCellBorders(new BottomBorder { Val = cellModel.BottomBorder.BorderValue.ToOOxml(), Color = cellModel.BottomBorder.Color, Size = Convert.ToUInt32(cellModel.BottomBorder.Size) }));

            if (cellModel.LeftBorder != null)
                tableCellProperties.Append(new TableCellBorders(new LeftBorder { Val = cellModel.LeftBorder.BorderValue.ToOOxml(), Color = cellModel.LeftBorder.Color, Size = Convert.ToUInt32(cellModel.LeftBorder.Size) }));

            if (cellModel.RightBorder != null)
                tableCellProperties.Append(new TableCellBorders(new RightBorder { Val = cellModel.RightBorder.BorderValue.ToOOxml(), Color = cellModel.RightBorder.Color, Size = Convert.ToUInt32(cellModel.RightBorder.Size) }));

            tc.Append(tableCellProperties);

            for (int i = 0; i < cellContents.Count; i++)
            {
                tc.Append(cellContents[i].ContentItem as Paragraph);
            }

            return new PlatformTableCell(tc);
        }

        public ITableCell CreateTableMergeCell(IRun run, TableCellPropertiesModel cellModel)
        {
            if (run == null)
                throw new ArgumentNullException("run must not be null");
            if (cellModel == null)
                throw new ArgumentNullException("cellModel must not be null");

            return CreateTableMergeCell(new List<IRun>() { run }, cellModel);
        }

        public ITableCell CreateTableMergeCell(IList<IRun> cellContents, TableCellPropertiesModel cellModel)
        {
            if(cellContents == null)
                throw new ArgumentNullException("cellContents must not be null");
            if (cellContents.Any(e => e == null))
                throw new ArgumentNullException("All elements of cellContents must be not null");
            if (cellModel == null)
                throw new ArgumentNullException("cellModel must not be null");

            TableCell tc = new TableCell();

            var tableCellProperties = new TableCellProperties();
            if (!string.IsNullOrWhiteSpace(cellModel.Width))
                if (cellModel.WidthUnit.OOxmlEquals(DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Nil))
                    tableCellProperties.Append(new TableCellWidth { Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Dxa, Width = cellModel.Width });
                else
                    tableCellProperties.Append(new TableCellWidth { Type = cellModel.WidthUnit.ToOOxml(), Width = cellModel.Width });
            else
                tableCellProperties.Append(new TableCellWidth { Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Auto });

            if (cellModel.Gridspan.HasValue)
                tableCellProperties.Append(new GridSpan { Val = cellModel.Gridspan.Value });

            if (cellModel.Shading != null)
                tableCellProperties.Append(cellModel.Shading.ContentItem as Shading);

            tableCellProperties.Append(new TableCellVerticalAlignment { Val = cellModel.TableVerticalAlignementValues.ToOOxml() });

            // Gestion de la fusion des cellules
            if (cellModel.Fusion)
            {
                if (cellModel.FusionChild)
                {
                    // Obligation de rajouter la balise vMerge, elle permet d'indiquer que la cellule doit fusionner avec la cellule précédente qui contient un vMerge w: val = Restart
                    tableCellProperties.Append(new VerticalMerge() { /*Val = MergedCellValues.Continue*/ });
                }
                else
                    tableCellProperties.Append(new VerticalMerge() { Val = MergedCellValues.Restart });
            }

            // Modification de la rotation du texte dans la cellule
            if (cellModel.TextDirectionValues.HasValue)
                tableCellProperties.Append(new TextDirection { Val = cellModel.TextDirectionValues.ToOOxml() });

            // Gestion des bordures de la cellule
            if (cellModel.TopBorder != null)
                tableCellProperties.Append(new TableCellBorders(new TopBorder { Val = cellModel.TopBorder.BorderValue.ToOOxml(), Color = cellModel.TopBorder.Color, Size = Convert.ToUInt32(cellModel.TopBorder.Size) }));

            if (cellModel.BottomBorder != null)
                tableCellProperties.Append(new TableCellBorders(new BottomBorder { Val = cellModel.BottomBorder.BorderValue.ToOOxml(), Color = cellModel.BottomBorder.Color, Size = Convert.ToUInt32(cellModel.BottomBorder.Size) }));

            if (cellModel.LeftBorder != null)
                tableCellProperties.Append(new TableCellBorders(new LeftBorder { Val = cellModel.LeftBorder.BorderValue.ToOOxml(), Color = cellModel.LeftBorder.Color, Size = Convert.ToUInt32(cellModel.LeftBorder.Size) }));

            if (cellModel.RightBorder != null)
                tableCellProperties.Append(new TableCellBorders(new RightBorder { Val = cellModel.RightBorder.BorderValue.ToOOxml(), Color = cellModel.RightBorder.Color, Size = Convert.ToUInt32(cellModel.RightBorder.Size) }));

            tc.Append(tableCellProperties);

            Paragraph par = new Paragraph();
            ParagraphProperties pr = new ParagraphProperties(new TableCellVerticalAlignment { Val = cellModel.TableVerticalAlignementValues.ToOOxml() });
            //new SpacingBetweenLines() { After = cellModel.SpacingAfter, Before = cellModel.SpacingBefore, Line = "240" });

            if (cellModel.Justification.HasValue)
                pr.Append(new Justification() { Val = cellModel.Justification.Value.ToOOxml() });
            par.Append(pr);

            if (cellContents.Count == 0)
                tc.Append(par);
            else
            {
                for (int i = 0; i < cellContents.Count; i++)
                {
                    par.AppendChild(cellContents[i].ContentItem as Paragraph);
                }
                tc.Append(par);
            }

            return new PlatformTableCell(tc);
        }

        public ITableRow CreateTableRow(IList<ITableCell> cells, TableRowPropertiesModel properties = null)
        {
            if (cells == null)
                throw new ArgumentNullException("cells must not be null");
            if (cells.Any(e => e == null))
                throw new ArgumentNullException("All elements of cells must be not null");

            var tr = new TableRow();

            var ptr = new PlatformTableRow(tr);

            if (properties != null)
            {
                // Create a TableRowProperties
                if (properties.Height.HasValue)
                {
                    TableRowProperties trPpr = new TableRowProperties(new TableRowHeight() { Val = Convert.ToUInt32(properties.Height) });
                    tr.Append(trPpr);
                }
            }

            if (cells != null)
            {
                ptr.Append(cells);
            }

            return ptr;
        }

        #endregion
    }
}
