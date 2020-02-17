using System;
using System.IO;
using it = iTextSharp.text;
using Pdf.Engine.ReportEngine.Renders;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using Pdf.Engine.ReportEngine.Extensions;
using iTextSharp.text.pdf;
using Pdf.Engine.ReportEngine;

namespace Pdf.Engine
{
    public class PdfManager : IDisposable
    {
        private FileStream output = null;

        public byte[] GenerateReport(Document document, ContextModel context, IFormatProvider formatProvider)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                var sizePage = (document.Pages[0] as Page).PageSize.ToRectangle();
                var margin = (document.Pages[0] as Page).Margin;
                var pdfDocument = new it.Document(sizePage,
                                                margin.Left / 20,
                                                margin.Right / 20,
                                                margin.Top / 20,
                                                margin.Bottom / 20);

                var writer = PdfWriter.GetInstance(pdfDocument, stream);
                writer.PageEmpty = true;

                var ctx = new EngineContext();

                try
                {
                    document.Render(writer, pdfDocument, context, ctx, formatProvider);
                }
                catch(Exception ex)
                {
                    int i = 0;
                }
                // Close the Document - this saves the document contents to the output stream
                if (pdfDocument.IsOpen())
                    pdfDocument.Close();
                writer.Close();
                return stream.ToArray();
            }
        }

        public void Dispose()
        {
            if (output != null)
                output.Dispose();
        }
    }
}
