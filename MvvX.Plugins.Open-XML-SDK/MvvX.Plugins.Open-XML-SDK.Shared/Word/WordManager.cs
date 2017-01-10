using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.Bookmarks;
using MvvX.Plugins.OpenXMLSDK.Word.Models;
using MvvX.Plugins.OpenXMLSDK.Word.Paragraphs;
using MvvX.Plugins.OpenXMLSDK.Word.Tables;
using MvvX.Plugins.OpenXMLSDK.Word.Tables.Models;
using MvvX.Plugins.OpenXMLSDK.Platform.Word.Bookmarks;
using MvvX.Plugins.OpenXMLSDK.Platform.Word.Extensions;
using MvvX.Plugins.OpenXMLSDK.Platform.Word.Paragraphs;
using MvvX.Plugins.OpenXMLSDK.Platform.Word.Tables;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using MvvX.Plugins.OpenXMLSDK.Word;
using MvvX.Plugins.OpenXMLSDK.Drawing.Pictures.Model;
using System.Text;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word
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
            memoryStream.Position = 0;
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

        /// <summary>
        /// Insert an html content on bookmark
        /// </summary>
        /// <param name="bookmark"></param>
        /// <param name="html"></param>
        public void SetHtmlOnBookmark(string bookmark, string html)
        {
            if (string.IsNullOrWhiteSpace(bookmark))
                throw new ArgumentNullException("bookmark must be not null or white spaces");
            if (wdDoc == null)
                throw new InvalidOperationException("Document not loaded");

            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(html)))
            {
                AddAltChunkOnBookmark(bookmark, ms, AlternativeFormatImportPartType.Xhtml);
            }
        }

        /// <summary>
        /// Insert another word document on bookmark
        /// </summary>
        /// <param name="bookmark"></param>
        /// <param name="content"></param>
        public void SetSubDocumentOnBookmark(string bookmark, Stream content)
        {
            if (string.IsNullOrWhiteSpace(bookmark))
                throw new ArgumentNullException("bookmark must be not null or white spaces");
            if (wdDoc == null)
                throw new InvalidOperationException("Document not loaded");

            AddAltChunkOnBookmark(bookmark, content, AlternativeFormatImportPartType.WordprocessingML);
        }

        /// <summary>
        /// Append SubDocument at end of current doc
        /// </summary>
        /// <param name="content"></param>
        public void AppendSubDocument(Stream content, bool withPageBreak)
        {
            if (wdDoc == null)
                throw new InvalidOperationException("Document not loaded");

            AlternativeFormatImportPart formatImportPart = wdMainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML);

            formatImportPart.FeedData(content);

            AltChunk altChunk = new AltChunk();
            altChunk.Id = wdMainDocumentPart.GetIdOfPart(formatImportPart);

            wdMainDocumentPart.Document.Body.Elements<Paragraph>().Last().InsertAfterSelf(altChunk);

            if (withPageBreak)
            {
                Paragraph p = new Paragraph(
                    new Run(
                        new Break() { Type = BreakValues.Page }));
                altChunk.InsertAfterSelf(p);
            }
        }

        /// <summary>
        /// Insert a document on bookmark
        /// </summary>
        /// <param name="bookmark"></param>
        /// <param name="content"></param>
        /// <param name="importType"></param>
        private void AddAltChunkOnBookmark(string bookmark, Stream content, AlternativeFormatImportPartType importType)
        {
            AlternativeFormatImportPart formatImportPart = wdMainDocumentPart.AddAlternativeFormatImportPart(importType);

            formatImportPart.FeedData(content);

            AltChunk altChunk = new AltChunk();
            altChunk.Id = wdMainDocumentPart.GetIdOfPart(formatImportPart);

            var bookmarkElement = FindBookmark(bookmark);
            if (bookmarkElement != null)
            {
                var paragraph = bookmarkElement.Ancestors<IParagraph>().LastOrDefault();
                // without an empty paragraph after altchunk, the docx might be corrupted if the bookmark is inside a table and the html only contains one paragraph
                if(paragraph.Ancestors<ITable>().Any())
                    paragraph.InsertAfterSelf(new PlatformParagraph());
                paragraph.InsertAfterSelf(new PlatformAltChunk(altChunk));
            }
        }
        #endregion

        #region Images

        /// <summary>
        /// Renvoie l'ID d'une part dans le document
        /// </summary>
        /// <typeparam name="T">Type du part</typeparam>
        /// <param name="part">Part</param>
        /// <returns>Id du part dans le document</returns>
        private string GetIdOfPart<T>(T part) where T : OpenXmlPart
        {
            return wdMainDocumentPart.GetIdOfPart(part);
        }

        /// <summary>
        /// Permet d'ajouter le type d'une image dans le document de type OpenXmlPart
        /// </summary>
        /// <param name="type">Type de l'image</param>
        /// <returns></returns>
        private ImagePart AddImagePart(ImagePartType type)
        {
            return wdMainDocumentPart.AddImagePart(type);
        }

        public IRun CreateImage(byte[] imageData, PictureModel model)
        {
            if (imageData == null)
                throw new ArgumentNullException("Image not found");

            ImagePart imagePart = AddImagePart((ImagePartType)(int)model.ImagePartType);

            using (MemoryStream stream = new MemoryStream(imageData))
            {
                imagePart.FeedData(stream);
            }

            return CreateImage(imagePart, model);
        }

        public IRun CreateImage(string fileName, PictureModel model)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentNullException("fileName must not be null or white space");
            if (!File.Exists(fileName))
                throw new ArgumentNullException("File not found");

            ImagePart imagePart = AddImagePart((ImagePartType)(int)model.ImagePartType);

            using (FileStream stream = new FileStream(fileName, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            return CreateImage(imagePart, model);
        }

        private IRun CreateImage(ImagePart imagePart, PictureModel model)
        {
            string relationshipId = GetIdOfPart(imagePart);

            long imageWidth = 990000L;
            long imageHeight = 792000L;

#if __WPF__ || __IOS__
            using (var bm = new System.Drawing.Bitmap(imagePart.GetStream()))
#endif
#if __ANDROID__
            using (var bm = Android.Graphics.BitmapFactory.DecodeStream(imagePart.GetStream()))
#endif
            {
                long bmWidth = (long)bm.Width;
                long bmHeight = (long)bm.Height;

                // Redimensionnement de l'image car trop grande en largeur:
                if (model.MaxWidth.HasValue && model.MaxWidth.Value < bmWidth)
                {
                    long ratio = model.MaxWidth.Value * 100L / bmWidth;

                    bmWidth = (long)((double)bmWidth * ((double)ratio / 100D));
                    bmHeight = (long)((double)bmHeight * ((double)ratio / 100D));
                }

                // Redimensionnement de l'image car trop grande en hauteur:
                if (model.MaxHeight.HasValue && model.MaxHeight.Value < bmHeight)
                {
                    long ratio = model.MaxHeight.Value * 100L / bmHeight;

                    bmWidth = (long)((double)bmWidth * ((double)ratio / 100D));
                    bmHeight = (long)((double)bmHeight * ((double)ratio / 100D));
                }

#if __WPF__ || __IOS__
                imageWidth = bmWidth * (long)((float)914400 / bm.HorizontalResolution);
                imageHeight = bmHeight * (long)((float)914400 / bm.VerticalResolution);
#endif
#if __ANDROID__
                // TODO : Check this method
                imageWidth = bmWidth * (long)((float)914400 / (long)bm.Density);
                imageHeight = bmHeight * (long)((float)914400 / (long)bm.Density);
#endif
            }

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

            result.Append(new DocumentFormat.OpenXml.Wordprocessing.Drawing(
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

#region Instanciate new SDK items

        public IRun CreateRun()
        {
            return Create<PlatformRun>();
        }

        public ITable CreateTable()
        {
            return Create<PlatformTable>();
        }

        private T Create<T>() where T : new()
        {
            return new T();
        }

        public int CreateBulletList()
        {
            NumberingDefinitionsPart numberingPart = wdMainDocumentPart.NumberingDefinitionsPart;
            if (numberingPart == null)
            {
                numberingPart = wdMainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                Numbering element = new Numbering();
                element.Save(numberingPart);
            }

            // Insert an AbstractNum into the numbering part numbering list.  The order seems to matter or it will not pass the 
            // Open XML SDK Productity Tools validation test.  AbstractNum comes first and then NumberingInstance and we want to
            // insert this AFTER the last AbstractNum and BEFORE the first NumberingInstance or we will get a validation error.
            var abstractNumberId = numberingPart.Numbering.Elements<AbstractNum>().Count() + 1;
            var lvls = new List<Level>();
            for (int i = 0; i < 6; i++)
            {
                var abstractLevel = new Level(new NumberingFormat() { Val = NumberFormatValues.Bullet }, new LevelText() { Val = "" }, new ParagraphProperties() { Indentation = new Indentation() { Left = (720 * (i+1)).ToString(), Hanging = "360" } }, new RunProperties() { RunFonts = new RunFonts() { Ascii = "Symbol", HighAnsi = "Symbol" } }) { LevelIndex = i };
                lvls.Add(abstractLevel);
            }
            var abstractNum1 = new AbstractNum(lvls) { AbstractNumberId = abstractNumberId };

            if (abstractNumberId == 1)
            {
                numberingPart.Numbering.Append(abstractNum1);
            }
            else
            {
                AbstractNum lastAbstractNum = numberingPart.Numbering.Elements<AbstractNum>().Last();
                numberingPart.Numbering.InsertAfter(abstractNum1, lastAbstractNum);
            }

            // Insert an NumberingInstance into the numbering part numbering list.  The order seems to matter or it will not pass the 
            // Open XML SDK Productity Tools validation test.  AbstractNum comes first and then NumberingInstance and we want to
            // insert this AFTER the last NumberingInstance and AFTER all the AbstractNum entries or we will get a validation error.
            var numberId = numberingPart.Numbering.Elements<NumberingInstance>().Count() + 1;
            NumberingInstance numberingInstance1 = new NumberingInstance() { NumberID = numberId };
            AbstractNumId abstractNumId1 = new AbstractNumId() { Val = abstractNumberId };
            numberingInstance1.Append(abstractNumId1);

            if (numberId == 1)
            {
                numberingPart.Numbering.Append(numberingInstance1);
            }
            else
            {
                var lastNumberingInstance = numberingPart.Numbering.Elements<NumberingInstance>().Last();
                numberingPart.Numbering.InsertAfter(numberingInstance1, lastNumberingInstance);
            }

            return numberId;
        }
        #endregion

        #region Texts

        public IRun CreateRunForTable(ITable table)
        {
            if (table == null)
                throw new ArgumentNullException("table must not be null");

            var run = new Run(table.ContentItem as Table);

            return new PlatformRun(run);
        }

        public IParagraph CreateParagraphForRun(IRun run, ParagraphPropertiesModel ppm = null)
        {
            if (run == null)
                throw new ArgumentNullException("run must not be null");

            var paragraph = new Paragraph();

            var platformParagraph = new PlatformParagraph(paragraph);

            if (ppm != null)
            {
                AutoMapper.Mapper.Map(ppm, platformParagraph.ParagraphProperties);
            }

            paragraph.Append(run.ContentItem as Run);

            return platformParagraph;
        }

        public IRun CreateRunForText(string content, RunPropertiesModel rpm = null)
        {
            var platformRun = new PlatformRun();

            if (rpm != null)
            {
                AutoMapper.Mapper.Map(rpm, platformRun.Properties);
            }

            platformRun.Append(PlatformText.New(content, OpenXMLSDK.Word.SpaceProcessingModeValues.Preserve));

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

            var pt = CreateTable();

            if (properties != null)
                AutoMapper.Mapper.Map(properties, pt.Properties);

            if (rows != null)
            {
                for (int i = 0; i < rows.Count; i++)
                {
                    pt.Append(rows[i]);
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

            var platformCellTable = PlatformTableCell.New();
            
            if (cellModel != null)
                AutoMapper.Mapper.Map(cellModel, platformCellTable.Properties);

            TableCell tc = platformCellTable.ContentItem as TableCell;

            var tableCellProperties = platformCellTable.Properties.ContentItem as TableCellProperties;
            
            //tableCellProperties.Append(new TableCellVerticalAlignment { Val = cellModel.TableVerticalAlignementValues.ToOOxml() });

            // Modification de la rotation du texte dans la cellule
            if (cellModel.TextDirectionValues.HasValue)
                tableCellProperties.Append(new TextDirection { Val = cellModel.TextDirectionValues.ToOOxml() });
            
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
            
            var platformCellTable = PlatformTableCell.New();

            if (cellModel != null)
                AutoMapper.Mapper.Map(cellModel, platformCellTable.Properties);

            TableCell tc = platformCellTable.ContentItem as TableCell;

            var tableCellProperties = platformCellTable.Properties.ContentItem as TableCellProperties;
            
            //tableCellProperties.Append(new TableCellVerticalAlignment { Val = cellModel.TableVerticalAlignementValues.ToOOxml() });

            // Modification de la rotation du texte dans la cellule
            if (cellModel.TextDirectionValues.HasValue)
                tableCellProperties.Append(new TextDirection { Val = cellModel.TextDirectionValues.ToOOxml() });
            
            for (int i = 0; i < cellContents.Count; i++)
            {
                platformCellTable.Append(cellContents[i]);
            }

            return platformCellTable;
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

            var platformCellTable = PlatformTableCell.New();

            if (cellModel != null)
                AutoMapper.Mapper.Map(cellModel, platformCellTable.Properties);

            TableCell tc = platformCellTable.ContentItem as TableCell;

            var tableCellProperties = platformCellTable.Properties.ContentItem as TableCellProperties;
            
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
                    par.AppendChild(cellContents[i].ContentItem as Run);
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
            
            var ptr = PlatformTableRow.New();

            if (properties != null)
            {
                AutoMapper.Mapper.Map(properties, ptr.Properties);
                // Create a TableRowProperties
                //if (properties.Height.HasValue)
                //{
                //    TableRowProperties trPpr = new TableRowProperties(new TableRowHeight() { Val = Convert.ToUInt32(properties.Height) });
                //    tr.Append(trPpr);
                //}
                //ptr.Properties.Height = 
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

