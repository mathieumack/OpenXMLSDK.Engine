using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word;
using MvvX.Open_XML_SDK.Core.Word.Bookmarks;
using MvvX.Open_XML_SDK.Core.Word.Images;
using MvvX.Open_XML_SDK.Core.Word.Paragraphs;
using MvvX.Open_XML_SDK.Core.Word.Tables;
using MvvX.Open_XML_SDK.Core.Word.Tables.Models;
using MvvX.Open_XML_SDK.Shared.Word;
using MvvX.Open_XML_SDK.Shared.Word.Bookmarks;
using MvvX.Open_XML_SDK.Shared.Word.Extensions;
using MvvX.Open_XML_SDK.Shared.Word.Paragraphs;
using MvvX.Open_XML_SDK.Shared.Word.Tables;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

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

            AutoMapperInitializer.Init();
        }

        #endregion

        #region Automapper

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

        #region Images

        public void InsertPictureToBookmark(string bookmark, string fileName, ImageType type, long? maxWidth = null, long? maxHeight = null)
        {
            if (string.IsNullOrWhiteSpace(bookmark))
                throw new ArgumentNullException("bookmark must be not null or white spaces");
            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentNullException("FileName must not be null or white space");
            if(!File.Exists(fileName))
                throw new FileNotFoundException("File not found");

            var enumType = (int)type;
            ImagePartType imageType = (ImagePartType)enumType;
            var image = CreateImage(fileName, imageType, maxWidth, maxHeight);

            var bookmarkElement = FindBookmark(bookmark);
            if (bookmarkElement != default(BookmarkEnd))
            {
                bookmarkElement.InsertAfterSelf(image);
            }
        }

        public IRun CreateImage(string fileName, ImagePartType type, long? maxWidth = null, long? maxHeight = null)
        {
            byte[] file = File.ReadAllBytes(fileName);
            MemoryStream ms = new MemoryStream(file);

            return CreateImage(ms, type, maxWidth, maxHeight);
        }

        public IRun CreateImage(Stream stream, ImagePartType type, long? maxWidth = null, long? maxHeight = null)
        {
            ImagePart imagePart = wdMainDocumentPart.AddImagePart(type);
            MemoryStream msClone = new MemoryStream();
            stream.CopyTo(msClone);
            stream.Position = 0;
            imagePart.FeedData(stream);

            string relationshipId = wdMainDocumentPart.GetIdOfPart(imagePart);

            return CreateImage(msClone, relationshipId, maxWidth, maxHeight);
        }
                
        private IRun CreateImage(Stream ms, string relationshipId, long? maxWidth = null, long? maxHeight = null)
        {
            long imageWidth = 990000L;
            long imageHeight = 792000L;
          
            var result = new Run();

            var runProperties = new RunProperties();
            runProperties.Append(new NoProof());
            result.Append(runProperties);

            var graphicFrameLocks = new A.GraphicFrameLocks() { NoChangeAspect = true };
            graphicFrameLocks.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            var picture = new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip()
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = imageWidth, Cy = imageHeight }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }));
            picture.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

            var graphic = new A.Graphic(
                             new A.GraphicData(
                                 picture
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" });
            graphic.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            result.Append(new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = imageWidth, Cy = imageHeight },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(graphicFrameLocks),
                         graphic
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     }));
            return new PlatformRun(result);
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
            var run = new Run();
            return new PlatformRun(run);
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
            var platformRun = PlatformRun.New();

            if (rpm != null)
            {
                AutoMapper.Mapper.Map(rpm, platformRun.Properties);
            }

            platformRun.Append(PlatformText.New(content, Core.Word.SpaceProcessingModeValues.Preserve));

            return platformRun;
        }

        #endregion

        #region Tables

        public ITable CreateTable(IList<ITableRow> rows, TablePropertiesModel properties)
        {
            if (rows == null)
                throw new ArgumentNullException("rows must not be null");
            if (rows.Any(e => e == null))
                throw new ArgumentNullException("All elements of rows must not be null");

            // Check on line cells numbers
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
            var pt = PlatformTable.New();
            var tbl = pt.ContentItem as Table;
            
            var tableProp = pt.Properties.ContentItem as TableProperties;

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

            if (properties != null)
                AutoMapper.Mapper.Map(properties, pt.Properties);

            if (rows != null)
            {
                for (int i = 0; i < rows.Count; i++)
                {
                    tbl.Append(rows[i].ContentItem as TableRow);
                }
            }

            return pt;
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

            //if (cellModel.Shading != null)
            //    tableCellProperties.Append(cellModel.Shading.ContentItem as Shading);

            //tableCellProperties.Append(new TableCellVerticalAlignment { Val = cellModel.TableVerticalAlignementValues.ToOOxml() });

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
            ParagraphProperties pr = new ParagraphProperties(); // new TableCellVerticalAlignment { Val = cellModel.TableVerticalAlignementValues.ToOOxml() });
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

            var platformTableCell = PlatformTableCell.New();

            var tableCellProperties = platformTableCell.Properties.ContentItem as TableCellProperties;
            if (!string.IsNullOrWhiteSpace(cellModel.Width))
                if (cellModel.WidthUnit.OOxmlEquals(DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Nil))
                    tableCellProperties.Append(new TableCellWidth { Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Dxa, Width = cellModel.Width });
                else
                    tableCellProperties.Append(new TableCellWidth { Type = cellModel.WidthUnit.ToOOxml(), Width = cellModel.Width });
            else
                tableCellProperties.Append(new TableCellWidth { Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Auto });

            if (cellModel.Gridspan.HasValue)
                platformTableCell.Properties.GridSpan.Val  = cellModel.Gridspan.Value;

            if (cellModel.Shading != null)
                tableCellProperties.Shading.Val = cellModel.Shading.ToOOxml();

            //tableCellProperties.Append(new TableCellVerticalAlignment { Val = cellModel.TableVerticalAlignementValues.ToOOxml() });

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
            
            for (int i = 0; i < cellContents.Count; i++)
            {
                platformTableCell.Append(cellContents[i]);
            }

            return platformTableCell;
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
            {
                tableCellProperties.GridSpan = new GridSpan();
                tableCellProperties.GridSpan.Val = cellModel.Gridspan.Value;
            }

            //if (cellModel.Shading != null)
            //    tableCellProperties.Shading.Val = cellModel.Shading.ToOOxml();

            //tableCellProperties.Append(new TableCellVerticalAlignment { Val = cellModel.TableVerticalAlignementValues.ToOOxml() });

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
            ParagraphProperties pr = new ParagraphProperties();// new TableCellVerticalAlignment { Val = cellModel.TableVerticalAlignementValues.ToOOxml() });
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
