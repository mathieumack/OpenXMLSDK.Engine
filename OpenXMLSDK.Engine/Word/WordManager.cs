using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.ReportEngine.DataContext;
using OpenXMLSDK.Engine.Word.ReportEngine;
using OpenXMLSDK.Engine.Word.ReportEngine.Renders;

namespace OpenXMLSDK.Engine.Word
{
    public class WordManager : IDisposable
    {
        #region Fields

        private MemoryStream streamFile;

        private string filePath;

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
            memoryStream.Position = 0;
            return memoryStream;
        }

        #endregion

        #region Settings

        /// <summary>
        /// Force or not the update of Table of Content when the document is opened
        /// </summary>
        /// <param name="updateToc"></param>
        public void SetToCUpdate(bool updateToc)
        {
            wdMainDocumentPart?.DocumentSettingsPart?.Settings?.Append(new UpdateFieldsOnOpen() { Val = new OnOffValue(updateToc) });
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
                throw new ArgumentNullException(nameof(filePath), "filePath must be not null or white spaces");
            if (!File.Exists(filePath))
                throw new FileNotFoundException("file not found", filePath);
            

            try
            {
                wdDoc = WordprocessingDocument.Open(filePath, isEditable);
                wdMainDocumentPart = wdDoc.MainDocumentPart;

                return true;
            }
            catch
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
                throw new ArgumentNullException(nameof(streamFile), "streamFile must be not null");

            try
            {
                wdDoc = WordprocessingDocument.Open(streamFile, isEditable);
                wdMainDocumentPart = wdDoc.MainDocumentPart;

                return true;
            }
            catch
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
                throw new ArgumentNullException(nameof(templateFilePath), "templateFilePath must be not null or white spaces");
            if (!File.Exists(templateFilePath))
                throw new FileNotFoundException("file not found");

            streamFile = new MemoryStream();
            try
            {
                byte[] byteArray = File.ReadAllBytes(templateFilePath);
                streamFile.Write(byteArray, 0, byteArray.Length);

                wdDoc = WordprocessingDocument.Open(streamFile, true);

                // Change the document type to Document
                wdDoc.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

                wdMainDocumentPart = wdDoc.MainDocumentPart;

                return true;
            }
            catch
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
                throw new ArgumentNullException(nameof(templateFilePath), "templateFilePath must be not null or white spaces");
            if (!File.Exists(templateFilePath))
                throw new FileNotFoundException("file not found");

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
            catch
            {
                wdDoc = null;
                return false;
            }
        }

        #endregion

        #region Bookmarks

        public BookmarkEnd FindBookmark(string bookmark)
        {
            var resultMain = wdMainDocumentPart.RootElement.Descendants<BookmarkStart>().FirstOrDefault(e => e.Name == bookmark);
            if (resultMain != default(BookmarkStart))
                return wdMainDocumentPart.RootElement.Descendants<BookmarkEnd>().FirstOrDefault(e => e.Id.Value == resultMain.Id);

            if (wdMainDocumentPart.HeaderParts != null)
            {
                foreach (var header in wdMainDocumentPart.HeaderParts)
                {
                    var result = header.RootElement.Descendants<BookmarkStart>().FirstOrDefault(e => e.Name == bookmark);
                    if (result != default(BookmarkStart))
                        return header.RootElement.Descendants<BookmarkEnd>().FirstOrDefault(e => e.Id.Value == result.Id);
                }
            }

            if (wdMainDocumentPart.FooterParts != null)
            {
                foreach (var footer in wdMainDocumentPart.FooterParts)
                {
                    var result = footer.RootElement.Descendants<BookmarkStart>().FirstOrDefault(e => e.Name == bookmark);
                    if (result != default(BookmarkStart))
                        return footer.RootElement.Descendants<BookmarkEnd>().FirstOrDefault(e => e.Id.Value == result.Id);
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

        public void SetOnBookmark(string bookmark, OpenXmlElement element)
        {
            if (string.IsNullOrWhiteSpace(bookmark))
                throw new ArgumentNullException(nameof(bookmark), "bookmark must be not null or white spaces");
            if (wdDoc == null)
                throw new InvalidOperationException("Document not loaded");

            var bookmarkElement = FindBookmark(bookmark);
            if (bookmarkElement != null)
                bookmarkElement.InsertAfterSelf(element);
        }

        public void SetParagraphsOnBookmark(string bookmark, IList<Paragraph> paragraphs)
        {
            if (string.IsNullOrWhiteSpace(bookmark))
                throw new ArgumentNullException(nameof(bookmark), "bookmark must be not null or white spaces");
            if (wdDoc == null)
                throw new InvalidOperationException("Document not loaded");

            var bookmarkElement = FindBookmark(bookmark);
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

        public void SetTextOnBookmark(string bookmark, string text)
        {
            if (string.IsNullOrWhiteSpace(bookmark))
                throw new ArgumentNullException(nameof(bookmark), "bookmark must be not null or white spaces");
            if (wdDoc == null)
                throw new InvalidOperationException("Document not loaded");

            var run = new Run(new Text(text));
            SetOnBookmark(bookmark, run);
        }

        public void SetTextsOnBookmark(string bookmark, List<string> texts, bool formated = true)
        {
            if (string.IsNullOrWhiteSpace(bookmark))
                throw new ArgumentNullException(nameof(bookmark), "bookmark must be not null or white spaces");
            if (wdDoc == null)
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

            SetOnBookmark(bookmark, run);
        }

        /// <summary>
        /// Insert an html content on bookmark
        /// </summary>
        /// <param name="bookmark"></param>
        /// <param name="html"></param>
        public void SetHtmlOnBookmark(string bookmark, string html)
        {
            if (string.IsNullOrWhiteSpace(bookmark))
                throw new ArgumentNullException(nameof(bookmark), "bookmark must be not null or white spaces");
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
                throw new ArgumentNullException(nameof(bookmark), "bookmark must be not null or white spaces");
            if (wdDoc == null)
                throw new InvalidOperationException("Document not loaded");

            AddAltChunkOnBookmark(bookmark, content, AlternativeFormatImportPartType.WordprocessingML);
        }

        /// <summary>
        /// Append SubDocument at end of current doc
        /// </summary>
        /// <param name="content"></param>
        /// <param name="withPageBreak">If true insert a page break before.</param>
        public void AppendSubDocument(Stream content, bool withPageBreak)
        {
            if (wdDoc == null)
                throw new InvalidOperationException("Document not loaded");

            AlternativeFormatImportPart formatImportPart = wdMainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML);
            formatImportPart.FeedData(content);

            AltChunk altChunk = new AltChunk();
            altChunk.Id = wdMainDocumentPart.GetIdOfPart(formatImportPart);

            OpenXmlElement lastElement = wdMainDocumentPart.Document.Body.LastChild;

            if(lastElement is SectionProperties)
            {
                lastElement.InsertBeforeSelf(altChunk);
                if (withPageBreak)
                {
                    SectionProperties sectionProps = (SectionProperties)lastElement.Clone();
                    var p = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
                    var ppr = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties();
                    p.AppendChild(ppr);
                    ppr.AppendChild(sectionProps);
                    altChunk.InsertBeforeSelf(p);
                }
            }
            else
            {
                lastElement.InsertAfterSelf(altChunk);
                if (withPageBreak)
                {
                    Paragraph p = new Paragraph(
                        new Run(
                            new Break() { Type = BreakValues.Page }));
                    altChunk.InsertBeforeSelf(p);
                }
            }
        }

        /// <summary>
        /// Append a list of SubDocuments at end of current doc
        /// </summary>
        /// <param name="filePath">Destination file path</param>
        /// <param name="filesToInsert">Documents to insert</param>
        /// <param name="insertPageBreaks">Indicate if a page break must be added before each document</param>
        public void AppendSubDocumentsList(string filePath, IList<MemoryStream> filesToInsert, bool withPageBreak)
        {
            wdDoc = WordprocessingDocument.Open(filePath, true);

            // Change the document type to Document
            wdDoc.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
            wdMainDocumentPart = wdDoc.MainDocumentPart;
            SaveDoc();
            CloseDocNoSave();

            //Append documents
            OpenDoc(filePath, true);
            foreach (var file in filesToInsert)
            {
                file.Position = 0;
                AppendSubDocument(file, withPageBreak);
            }

            SaveDoc();
            CloseDocNoSave();
        }

        /// <summary>
        ///  Append a list of SubDocuments after the predecessorElement.
        /// </summary>
        /// <param name="filesToInsert"></param>
        /// <param name="insertPageBreaks"></param>
        /// <param name="predecessorElement"></param>
        private void AppendStreams(IList<MemoryStream> filesToInsert, bool insertPageBreaks, OpenXmlCompositeElement insertAfterElement)
        {
            OpenXmlCompositeElement openXmlCompositeElement = null;
            foreach (var file in filesToInsert)
            {
                using (WordprocessingDocument pkgSourceDoc = WordprocessingDocument.Open(file, true))
                {
                    var headers = pkgSourceDoc.MainDocumentPart.Document.Descendants<HeaderReference>().ToList();
                    foreach (var header in headers)
                        header.Remove();
                    var footers = pkgSourceDoc.MainDocumentPart.Document.Descendants<FooterReference>().ToList();
                    foreach (var footer in footers)
                        footer.Remove();
                    pkgSourceDoc.MainDocumentPart.Document.Save();
                }

                string altChunkId = "AltChunkId-" + Guid.NewGuid();

                AlternativeFormatImportPart chunk = wdMainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML, altChunkId);
                file.Position = 0;
                chunk.FeedData(file);

                AltChunk altChunk = new AltChunk() { Id = altChunkId };

                if (openXmlCompositeElement == null)
                    openXmlCompositeElement = insertAfterElement;

                if (insertPageBreaks)
                {
                    var run = new Run(new Break() { Type = BreakValues.Page });
                    Paragraph paragraph = new Paragraph(run);
                    openXmlCompositeElement.InsertAfterSelf<Paragraph>(paragraph);
                    openXmlCompositeElement = paragraph;
                }
                openXmlCompositeElement.InsertAfterSelf<AltChunk>(altChunk);
                openXmlCompositeElement = altChunk;

                if (wdDoc == null)
                    throw new InvalidOperationException("Document not loaded");
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
                var paragraph = bookmarkElement.Ancestors<Paragraph>().LastOrDefault();
                // without an empty paragraph after altchunk, the docx might be corrupted if the bookmark is inside a table and the html only contains one paragraph
                if (paragraph.Ancestors<Table>().Any())
                    paragraph.InsertAfterSelf(new Paragraph());
                paragraph.InsertAfterSelf(new AltChunk(altChunk));
            }
        }

        /// <summary>
        /// Allow to merge documents to insert it in a bookmark
        /// </summary>
        /// <param name="bookmark">Bookmark name</param>
        /// <param name="filesToInsert">Documents to insert</param>
        /// <param name="insertPageBreaks">Indicate if a page break must be added after each document</param>
        public void InsertDocsToBookmark(string bookmark, IList<MemoryStream> filesToInsert, bool insertPageBreaks)
        {
            if (string.IsNullOrWhiteSpace(bookmark))
                throw new ArgumentNullException(nameof(bookmark), "Bookmark must not be null or white space");
            if (filesToInsert == null)
                throw new ArgumentNullException(nameof(filesToInsert), "FilesToInsert must not be null");

            var bookmarkElement = FindBookmark(bookmark);
            if (bookmarkElement != default(BookmarkEnd))
            {
                OpenXmlCompositeElement insertAfterElement = wdDoc.MainDocumentPart.Document.Body.Descendants<BookmarkStart>().SingleOrDefault<BookmarkStart>((BookmarkStart b) => b.Name == bookmark)
                    .Ancestors<Paragraph>().FirstOrDefault<Paragraph>();
                AppendStreams(filesToInsert, insertPageBreaks, insertAfterElement);
            }
        }

        #endregion

        #region Parts

        /// <summary>
        /// Add a new part in the document
        /// </summary>
        /// <typeparam name="T">Type of the object</typeparam>
        /// <returns>created part</returns>
        public T AddNewPart<T>() where T : OpenXmlPart, IFixedContentTypePart
        {
            return wdMainDocumentPart.AddNewPart<T>();
        }

        /// <summary>
        /// Renvoie l'ID d'une part dans le document
        /// </summary>
        /// <typeparam name="T">Type du part</typeparam>
        /// <param name="part">Part</param>
        /// <returns>Id du part dans le document</returns>
        public string GetIdOfPart<T>(T part) where T : OpenXmlPart
        {
            return wdMainDocumentPart.GetIdOfPart(part);
        }

        #endregion

        #region Report Engine
        
        public byte[] GenerateReport(ReportEngine.Models.Document document, ContextModel context, IFormatProvider formatProvider)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                wdDoc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
                wdDoc.AddMainDocumentPart();
                wdDoc.MainDocumentPart.Document = new Document(new Body());

                document.Render(wdDoc, context, formatProvider);

                wdDoc.MainDocumentPart.Document.Save();
                wdDoc.Close();
                return stream.ToArray();
            }
        }

        /// <summary>
        /// Generate a word document 
        /// </summary>
        /// <param name="reportList">A list of Reports</param>
        /// <param name="mergeStyles">Indicates whether or not styles are merged</param>
        /// <returns></returns>
        public byte[] GenerateReport(IList<Report> reportList, bool mergeStyles, IFormatProvider formatProvider)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                wdDoc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
                wdDoc.AddMainDocumentPart();
                wdDoc.MainDocumentPart.Document = new Document(new Body());
                // add styles in document
                var spart = wdDoc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                spart.Styles = new Styles();
                IList<string> stylesId = null;
                if (mergeStyles)
                {
                    stylesId = reportList.Where(report => report.Document.Styles != null)
                                            .SelectMany(r => r.Document.Styles)
                                            .Select(style => style.StyleId)?.Distinct()?.ToList();                    
                }

                foreach (Report report in reportList)
                {
                    if (report.Document.Styles != null)
                    {
                        foreach (var style in report.Document.Styles)
                        {
                            if(!mergeStyles)
                            {
                                style.Render(spart, report.ContextModel);
                            }
                            else if (stylesId != null && stylesId.Count > 0 && stylesId.Contains(style.StyleId))
                            {
                                stylesId.Remove(style.StyleId);
                                style.Render(spart, report.ContextModel);
                            }
                        }
                    }

                    // Document render
                    report.Document.Render(wdDoc, report.ContextModel, report.AddPageBreak, formatProvider);

                    // footers
                    foreach (var footer in report.Document.Footers)
                    {
                        footer.Render(report.Document, wdDoc.MainDocumentPart, report.ContextModel, formatProvider);
                    }
                    // headers
                    foreach (var header in report.Document.Headers)
                    {
                        header.Render(report.Document, wdDoc.MainDocumentPart, report.ContextModel, formatProvider);
                    }
                }

                //Replace Last page Break
                if (wdDoc.MainDocumentPart.Document.Body.LastChild != null &&
                    wdDoc.MainDocumentPart.Document.Body.LastChild is Paragraph &&
                    wdDoc.MainDocumentPart.Document.Body.LastChild.FirstChild != null &&
                    wdDoc.MainDocumentPart.Document.Body.LastChild.FirstChild is ParagraphProperties &&
                    wdDoc.MainDocumentPart.Document.Body.LastChild.FirstChild.FirstChild != null &&
                    wdDoc.MainDocumentPart.Document.Body.LastChild.FirstChild.FirstChild is SectionProperties)
                {
                    Paragraph lastChild = (Paragraph)wdDoc.MainDocumentPart.Document.Body.LastChild;
                    SectionProperties sectionPropertie = (SectionProperties)lastChild.FirstChild.FirstChild.Clone();
                    wdDoc.MainDocumentPart.Document.Body.ReplaceChild(sectionPropertie, wdDoc.MainDocumentPart.Document.Body.LastChild);
                }

                wdDoc.MainDocumentPart.Document.Save();
                wdDoc.Close();
                return stream.ToArray();
            }
        }

        #endregion
    }
}