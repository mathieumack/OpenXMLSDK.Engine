using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word;
using MvvX.Open_XML_SDK.Core.Word.Bases;
using MvvX.Open_XML_SDK.Core.Word.Bookmarks;
using MvvX.Open_XML_SDK.Core.Word.Images;
using MvvX.Open_XML_SDK.Core.Word.Paragraphs;
using MvvX.Open_XML_SDK.Core.Word.Tables;
using MvvX.Open_XML_SDK.Core.Word.Tables.Models;
using MvvX.Open_XML_SDK.Shared.Word.Tables;
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

        #region Images

        public void InsertPictureToBookmark(string bookmark, string fileName, ImageType type, long? maxWidth = null, long? maxHeight = null)
        {
            if (string.IsNullOrWhiteSpace(bookmark))
                throw new ArgumentNullException("bookmark must be not null or white spaces");
            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentNullException("FileName must not be null or white space");
            if(!File.Exists(fileName))
                throw new FileNotFoundException("File not found");

            var image = CreateImage(fileName, type, maxWidth, maxHeight);

            var bookmarkElement = FindBookmark(bookmark);
            if (bookmarkElement != default(BookmarkEnd))
            {
                bookmarkElement.InsertAfterSelf(image);
            }
        }
        public IRun CreateImage(string fileName, ImageType type, long? maxWidth = null, long? maxHeight = null)
        {
            byte[] file = File.ReadAllBytes(fileName);
            MemoryStream ms = new MemoryStream(file);

            return CreateImage(ms, type, maxWidth, maxHeight);
        }

        public IRun CreateImage(Stream stream, ImageType type, long? maxWidth = null, long? maxHeight = null)
        {

            var enumType = (int)type;
            ImagePartType imageType = (ImagePartType)enumType;
            ImagePart imagePart = wdMainDocumentPart.AddImagePart(imageType);
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

        #region Texts

        public IRun CreateTexte(string content, RunPropertiesModel rpm = null)
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

        public ITable CreateTable(IList<ITableRow> rows, TablePropertiesModel properties = null)
        {
            if (rows == null)
                throw new ArgumentNullException("row must be not null");
            
            Table table = new Table();
            PlatformTable tbl = new PlatformTable(table);

            TableProperties tblProps = new TableProperties();
            PlatformTableProperties tableProp = new PlatformTableProperties(tblProps);

            TableStyle tblStyle = new TableStyle() { Val = "TableGrid" };
            PlatformTableStyle tableStyle = new PlatformTableStyle(tblStyle);

            TableWidth tblWidth = new TableWidth();

            if (properties != null && !string.IsNullOrWhiteSpace(properties.Width))
            {
                tblWidth.Width = properties.Width;
                if (properties.WidthUnit.Equals(TableWidthUnitValues.Nil))
                    tblWidth.Type = TableWidthUnitValues.Pct;
                else
                {
                    var enumWidth = (int)properties.WidthUnit;
                    TableWidthUnitValues width = (TableWidthUnitValues)enumWidth;
                    tblWidth.Type = width;
                }

            }
            else
            {
                tblWidth.Width = "5000";
                tblWidth.Type = TableWidthUnitValues.Pct;
            }

            PlatformTableWidth tableWidth = new PlatformTableWidth(tblWidth);

            tableProp.Append(tableStyle, tableWidth);
            tbl.AppendChild(tableProp);

            if (properties != null)
            {
                //Convert IOpenXml enum to the right enum type
                var enumTop = (int)properties.TopBorder.BorderValue;
                BorderValues topBorder = (BorderValues)enumTop;
                var enumBot = (int)properties.BottomBorder.BorderValue;
                BorderValues botBorder = (BorderValues)enumBot;
                var enumLeft = (int)properties.LeftBorder.BorderValue;
                BorderValues leftBorder = (BorderValues)enumLeft;
                var enumRight = (int)properties.RightBorder.BorderValue;
                BorderValues rightBorder = (BorderValues)enumRight;
                var enumHor = (int)properties.InsideHorizontalBorder.BorderValue;
                BorderValues horBorder = (BorderValues)enumHor;
                var enumVer = (int)properties.InsideVerticalBorder.BorderValue;
                BorderValues verBorder = (BorderValues)enumVer;
                var enumJustif = (int)properties.TableJustification;
                TableRowAlignmentValues justif = (TableRowAlignmentValues)enumJustif;
                // Create a PlatformTableProperties object and specify its border information.
                PlatformTableProperties tblProp = new PlatformTableProperties(
                    new TableBorders(
                        new TopBorder() { Val = new EnumValue<BorderValues>(topBorder), Size = properties.TopBorder.Size, Color = properties.TopBorder.Color },
                        new BottomBorder() { Val = new EnumValue<BorderValues>(botBorder), Size = properties.BottomBorder.Size, Color = properties.BottomBorder.Color },
                        new LeftBorder() { Val = new EnumValue<BorderValues>(leftBorder), Size = properties.LeftBorder.Size, Color = properties.LeftBorder.Color },
                        new RightBorder() { Val = new EnumValue<BorderValues>(rightBorder), Size = properties.RightBorder.Size, Color = properties.RightBorder.Color },
                        new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(horBorder), Size = properties.InsideHorizontalBorder.Size, Color = properties.InsideHorizontalBorder.Color },
                        new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(verBorder), Size = properties.InsideVerticalBorder.Size, Color = properties.InsideVerticalBorder.Color }
                    ),
                    new TableProperties(
                        new TableJustification() { Val = justif }
                    )
                );
                // Append the TableProperties object to the empty table.
                tbl.AppendChild(tblProp);
            }


            if (rows != null)
            {
                for (int i = 0; i < rows.Count; i++)
                {
                    tbl.Append(rows[i]);
                }
            }

            return tbl;
        }

        public ITableCell CreateTableCell(IRun cellContent, TableCellPropertiesModel cellModel)
        {
            if (cellContent == null)
                throw new ArgumentNullException("cellContent must be not null");
            else if (cellModel == null)
                throw new ArgumentNullException("cellModel must be not null");

            return CreateTableCell(new List<IRun>() { cellContent }, cellModel);
        }

        public ITableCell CreateTableCell(IList<IRun> cellContents, TableCellPropertiesModel cellModel)
        {
            if (cellContents == null)
                throw new ArgumentNullException("cellContent must be not null");
            else if (cellModel == null)
                throw new ArgumentNullException("cellModel must be not null");

            TableCell tc = new TableCell();

            var tableCellProperties = new TableCellProperties();
            if (!string.IsNullOrWhiteSpace(cellModel.Width))
                if (cellModel.WidthUnit.Equals(TableWidthUnitValues.Nil))
                    tableCellProperties.Append(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = cellModel.Width });
                else
                {
                    var enumWidth = (int)cellModel.WidthUnit;
                    TableWidthUnitValues width = (TableWidthUnitValues)enumWidth;
                    tableCellProperties.Append(new TableCellWidth { Type = width, Width = cellModel.Width });
                }
            else
                tableCellProperties.Append(new TableCellWidth { Type = TableWidthUnitValues.Auto });

            if (cellModel.Gridspan.HasValue)
                tableCellProperties.Append(new GridSpan { Val = cellModel.Gridspan.Value });

            //TODO Shading
            //if (cellModel.Shading != null)
            //    tableCellProperties.Append(cellModel.Shading);


            var enumType = (int)cellModel.TableVerticalAlignementValues;
            TableVerticalAlignmentValues verticalDirection = (TableVerticalAlignmentValues)enumType;
            tableCellProperties.Append(new TableCellVerticalAlignment { Val = verticalDirection });

            // Text rotation
            if (cellModel.TextDirectionValues.HasValue)
            {
                enumType = (int)cellModel.TextDirectionValues;
                TextDirectionValues directionValue = (TextDirectionValues)enumType;

                tableCellProperties.Append(new TextDirection { Val = directionValue });
            }

            tc.Append(tableCellProperties);

            Paragraph par = new Paragraph();
            ParagraphProperties pr = new ParagraphProperties(new TableCellVerticalAlignment { Val = verticalDirection });

            if (cellModel.Justification.HasValue)
            {
                var enumJustif = (int)cellModel.Justification;
                JustificationValues justificationValue = (JustificationValues)enumJustif;

                pr.Append(new Justification() { Val = justificationValue });
            }
            par.Append(pr);

            PlatformParagraph para = new PlatformParagraph(par);
            for (int i = 0; i < cellContents.Count; i++)
                para.Append(cellContents[i]);
            
            tc.Append(par);
            return new PlatformTableCell(tc);
        }

        public ITableCell CreateTableCell(IList<IParagraph> cellContents, TableCellPropertiesModel cellModel)
        {
            if (cellContents == null)
                throw new ArgumentNullException("cellContents must be not null");
            else if (cellModel == null)
                throw new ArgumentNullException("cellModel must be not null");

            TableCell tc = new TableCell();

            var tableCellProperties = new TableCellProperties();
            if (!string.IsNullOrWhiteSpace(cellModel.Width))
                if (cellModel.WidthUnit.Equals(TableWidthUnitValues.Nil))
                    tableCellProperties.Append(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = cellModel.Width });
                else
                {
                    var enumWidth = (int)cellModel.WidthUnit;
                    TableWidthUnitValues width = (TableWidthUnitValues)enumWidth;
                    tableCellProperties.Append(new TableCellWidth { Type = width, Width = cellModel.Width });

                }
            else
                tableCellProperties.Append(new TableCellWidth { Type = TableWidthUnitValues.Auto });

            if (cellModel.Gridspan.HasValue)
                tableCellProperties.Append(new GridSpan { Val = cellModel.Gridspan.Value });

            //TODO Shading
            //if (cellModel.Shading != null)
            //    tableCellProperties.Append(cellModel.Shading);



            var enumDirection = (int)cellModel.TableVerticalAlignementValues;
            TableVerticalAlignmentValues directionValue = (TableVerticalAlignmentValues)enumDirection;
            tableCellProperties.Append(new TableCellVerticalAlignment { Val = directionValue });

            // Text rotation
            if (cellModel.TextDirectionValues.HasValue)
            {
                var enumText = (int)cellModel.TextDirectionValues;
                TextDirectionValues textDirection = (TextDirectionValues)enumText;

                tableCellProperties.Append(new TextDirection { Val = textDirection });
            }

            tc.Append(tableCellProperties);
            PlatformTableCell tcell = new PlatformTableCell(tc);
            for (int i = 0; i < cellContents.Count; i++)
            {
                tcell.Append(cellContents[i]);
            }

            return tcell;
        }

        public ITableCell CreateTableMergeCell(IRun cellContent, TableCellPropertiesModel cellModel)
        {
            if (cellContent == null)
                throw new ArgumentNullException("cellContent must be not null");
            else if (cellModel == null)
                throw new ArgumentNullException("cellModel must be not null");

            return CreateTableMergeCell(new List<IRun>() { cellContent }, cellModel);
        }

        public ITableCell CreateTableMergeCell(IList<IRun> cellContents, TableCellPropertiesModel cellModel)
        {
            if (cellContents == null)
                throw new ArgumentNullException("cellContent must be not null");
            else if (cellModel == null)
                throw new ArgumentNullException("cellModel must be not null");

            TableCell tc = new TableCell();

            var tableCellProperties = new TableCellProperties();
            if (!string.IsNullOrWhiteSpace(cellModel.Width))
                if (cellModel.WidthUnit.Equals(TableWidthUnitValues.Nil))
                    tableCellProperties.Append(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = cellModel.Width });
                else
                {
                    var enumType = (int)cellModel.WidthUnit;
                    TableWidthUnitValues width = (TableWidthUnitValues)enumType;
                    tableCellProperties.Append(new TableCellWidth { Type = width, Width = cellModel.Width });

                }
            else
                tableCellProperties.Append(new TableCellWidth { Type = TableWidthUnitValues.Auto });

            //TODO Shading
            //if (cellModel.Shading != null)
            //    tableCellProperties.Append(cellModel.Shading);


            var enumDirection = (int)cellModel.TableVerticalAlignementValues;
            TableVerticalAlignmentValues directionValue = (TableVerticalAlignmentValues)enumDirection;
            tableCellProperties.Append(new TableCellVerticalAlignment { Val = directionValue });

            // Cell fusion
            if (cellModel.Fusion)
            {
                tableCellProperties.Append(new VerticalMerge() { Val = MergedCellValues.Restart });
            }

            // Text rotation
            if (cellModel.TextDirectionValues.HasValue)
            {
                var enumText = (int)cellModel.TextDirectionValues;
                TextDirectionValues textDirection = (TextDirectionValues)enumText;

                tableCellProperties.Append(new TextDirection { Val = textDirection });
            }

            tc.Append(tableCellProperties);

            Paragraph par = new Paragraph();
            ParagraphProperties pr = new ParagraphProperties(new TableCellVerticalAlignment { Val = directionValue });

            if (cellModel.Justification.HasValue)
            {
                var enumText = (int)cellModel.Justification.Value;
                JustificationValues justificationValue = (JustificationValues)enumText;
                pr.Append(new Justification() { Val = justificationValue });
            }
            par.Append(pr);

            if (cellContents.Count == 0)
                tc.Append(par);
            else
            {
                PlatformParagraph para = new PlatformParagraph(par);
                for (int i = 0; i < cellContents.Count; i++)
                {
                    para.AppendChild(cellContents[i]);
                }
                tc.Append(par);
            }

            return new PlatformTableCell(tc);
        }

        public ITableRow CreateTableRow(IList<ITableCell> cells, TableRowPropertiesModel properties = null)
        {
            if (cells == null)
                throw new ArgumentNullException("cells must be not null");

            TableRow tra = new TableRow();
            var tr = new PlatformTableRow(tra);
            if (properties != null)
            {
                // Create a TableRowProperties
                if (properties.Height != 0)
                {
                    var height = Convert.ToUInt32(properties.Height);
                    TableRowProperties trPpr = new TableRowProperties(new TableRowHeight() { Val = height });
                    PlatformTableRowProperties ptrp = new PlatformTableRowProperties(trPpr);
                    tr.Append(ptrp);
                }
            }
            if (cells != null)
            {
                tr.Append(cells);
            }

            return tr;
        }

        #endregion

        #region Basics

        public IRun CreateRun(object contentItem)
        {
            return new PlatformRun() { ContentItem = contentItem };
        }
        public IParagraph CreateParagraph(object contentItem)
        {
            return new PlatformParagraph() { ContentItem = contentItem };
        }

        #endregion
    }
}
