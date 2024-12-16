using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
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
        /// Current Word object
        /// </summary>
        private WordprocessingDocument wdDoc = null;

        /// <summary>
        /// Current Word object
        /// </summary>
        internal WordprocessingDocument WordprocessingDocument
        {
            get { return wdDoc; }
        }

        /// <summary>
        /// Main page
        /// </summary>
        private MainDocumentPart wdMainDocumentPart = null;

        /// <summary>
        /// Main page
        /// </summary>
        internal MainDocumentPart MainDocumentPart
        {
            get { return wdMainDocumentPart; }
        }

        #endregion

        #region Constructors

        /// <summary>
        /// 
        /// </summary>
        public WordManager()
        {
            filePath = string.Empty;
        }

        #endregion

        #region Dispose / End

        public void CloseDoc()
        {
            this.SaveDoc();
            wdDoc.Dispose();
        }

        public void Dispose()
        {
            if (wdDoc != null)
                wdDoc.Dispose();
        }

        /// <summary>
        /// Allows to return the MemoryStream associated with the current document
        /// </summary>
        /// <returns>Current MemoryStream, null otherwise</returns>
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

        #region Private Functions - Open/Save/Create

        /// <summary>
        /// Save the current document
        /// </summary>
        public void SaveDoc()
        {
            if (wdDoc == null)
                throw new InvalidOperationException("Document not loaded");

            wdDoc.Save();
        }

        /// <summary>
        /// Close the document with automatic save
        /// </summary>
        public void CloseDocNoSave()
        {
            if (wdDoc == null)
                throw new InvalidOperationException("Document not loaded");

            wdDoc.Dispose();
        }

        /// <summary>
        /// Open a document
        /// </summary>
        /// <param name="filePath">Path and full name of the file to open</param>
        /// <param name="isEditable">Indicates if the file should be opened in editable mode (Read/Write)</param>
        /// <returns>True if the document was successfully opened</returns>
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
        /// Create a new empty document
        /// </summary>
        /// <returns></returns>
        public bool New()
        {
            streamFile = new MemoryStream();
            try
            {
                wdDoc = WordprocessingDocument.Create(streamFile, WordprocessingDocumentType.Document);
                wdMainDocumentPart = wdDoc.AddMainDocumentPart();
                wdMainDocumentPart.Document = new Document();

                var body = new Body();
                wdMainDocumentPart.Document.Append(body);

                return true;
            }
            catch
            {
                wdDoc = null;
                return false;
            }
        }

        /// <summary>
        /// Open a document in streaming mode
        /// </summary>
        /// <param name="streamFile">File stream</param>
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
        /// Open a document from a dotx template
        /// </summary>
        /// <param name="templateFileStream">Path and full name of the template</param>
        /// <returns>True if the document was successfully opened</returns>
        public bool OpenDocFromTemplate(Stream templateFileStream)
        {
            if (templateFileStream is null || templateFileStream == Stream.Null)
                throw new ArgumentNullException(nameof(templateFileStream), "templateFilePath must not be null");

            streamFile = new MemoryStream();
            try
            {
                var tempFilePath = Path.Combine(Environment.CurrentDirectory, Path.GetRandomFileName());
                using (var file = File.Create(tempFilePath))
                {
                    templateFileStream.Position = 0;
                    templateFileStream.CopyTo(file);
                }

                var templateWdDoc = WordprocessingDocument.CreateFromTemplate(tempFilePath);

                wdDoc = templateWdDoc.Clone(streamFile);

                wdMainDocumentPart = wdDoc.MainDocumentPart;

                File.Delete(tempFilePath);

                return true;
            }
            catch
            {
                wdDoc = null;
                return false;
            }
        }

        /// <summary>
        /// Open a document from a dotx template
        /// </summary>
        /// <param name="templateFilePath">Path and full name of the template</param>
        /// <returns>True if the document was successfully opened</returns>
        public bool OpenDocFromTemplate(string templateFilePath)
        {
            if (string.IsNullOrWhiteSpace(templateFilePath))
                throw new ArgumentNullException(nameof(templateFilePath), "templateFilePath must be not null or white spaces");
            if (!File.Exists(templateFilePath))
                throw new FileNotFoundException("file not found");

            streamFile = new MemoryStream();
            try
            {
                var templateWdDoc = WordprocessingDocument.CreateFromTemplate(templateFilePath);

                wdDoc = templateWdDoc.Clone(streamFile);

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
        /// Open a document from a dotx template
        /// </summary>
        /// <param name="templateFilePath">Path and full name of the template</param>
        /// <param name="newFilePath">Path and full name of the file to be saved</param>
        /// <param name="isEditable">Indicates if the file should be opened in editable mode (Read/Write)</param>
        /// <returns>True if the document was successfully opened</returns>
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

        #region Appends

        public void AppendSubDocument(List<Report> reportList, bool mergeStyles, CultureInfo formatProvider)
        {
            // add styles in document
            StyleDefinitionsPart spart = wdDoc.MainDocumentPart.StyleDefinitionsPart;
            IList<string> stylesId = null;
            if (wdDoc.MainDocumentPart.StyleDefinitionsPart == null)
            {
                spart = wdDoc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                spart.Styles = new Styles();
                if (mergeStyles)
                {
                    stylesId = reportList.Where(report => report.Document.Styles != null)
                                        .SelectMany(r => r.Document.Styles)
                                        .Select(style => style.StyleId)?.Distinct()?.ToList();
                }
            }

            foreach (Report report in reportList)
            {
                if (report.Document.Styles != null)
                {
                    foreach (var style in report.Document.Styles)
                    {
                        if (!mergeStyles)
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

            if (lastElement is SectionProperties)
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
        internal void AppendStreams(IList<MemoryStream> filesToInsert, bool insertPageBreaks, OpenXmlCompositeElement insertAfterElement)
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

        #endregion
    }
}
