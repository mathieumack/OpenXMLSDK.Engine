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
    }
}
