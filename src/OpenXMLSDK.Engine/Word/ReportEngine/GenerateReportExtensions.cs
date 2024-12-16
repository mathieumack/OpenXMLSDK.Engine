using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.ReportEngine.DataContext;
using OpenXMLSDK.Engine.Word.ReportEngine.Renders;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OpenXMLSDK.Engine.Word.ReportEngine
{
    public static class GenerateReportExtensions
    {

        #region Report Engine

        public static byte[] GenerateReport(this WordManager wordManager, ReportEngine.Models.Document document, ContextModel context, IFormatProvider formatProvider)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                var wdDoc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
                wdDoc.AddMainDocumentPart();
                wdDoc.MainDocumentPart.Document = new Document(new Body());

                document.Render(wdDoc, context, formatProvider);

                wdDoc.MainDocumentPart.Document.Save();
                wdDoc.Dispose();
                return stream.ToArray();
            }
        }

        /// <summary>
        /// Generate a word document 
        /// </summary>
        /// <param name="reportList">A list of Reports</param>
        /// <param name="mergeStyles">Indicates whether or not styles are merged</param>
        /// <returns></returns>
        public static byte[] GenerateReport(this WordManager wordManager, IList<Report> reportList, bool mergeStyles, IFormatProvider formatProvider)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                var wdDoc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
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

                wdDoc.MainDocumentPart.Document.Save();
                wdDoc.Dispose();
                return stream.ToArray();
            }
        }

        #endregion
    }
}
