using RCT = ReportEngine.Core.Template;
using RCD = ReportEngine.Core.DataContext;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Collections.Generic;

namespace Pdf.Engine.ReportEngine.Helpers
{
    /// <summary>
    /// Event used to create footer and header on each pages
    /// </summary>
    internal class PageEventHelper : PdfPageEventHelper
    {
        private readonly IList<RCT.Footer> footerModel;
        private readonly IList<RCT.Header> headerModel;

        private readonly EngineContext ctx;
        private readonly RCD.ContextModel parameters;

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="footer">Footer model</param>
        public PageEventHelper(IList<RCT.Header> header, IList<RCT.Footer> footer, EngineContext ctx, RCD.ContextModel parameters)
        {
            headerModel = header;
            footerModel = footer;
            this.ctx = ctx;
            this.parameters = parameters;
        }

        ///// <summary>
        ///// Event raised when the document is loaded
        ///// </summary>
        ///// <param name="writer"></param>
        ///// <param name="document"></param>
        //public override void OnOpenDocument(PdfWriter writer, Document document)
        //{
        //    footerModel.OnOpenDocumentRender(writer, document);
        //}

        ///// <summary>
        ///// Event raised when creating a new page
        ///// Foreach page
        ///// </summary>
        ///// <param name="writer"></param>
        ///// <param name="document"></param>
        //public override void OnEndPage(PdfWriter writer, Document document)
        //{
        //    footerModel.OnEndPageRender(writer, document, ctx, parameters);

        //    // On rajoute le header :
        //    if (headerModel != null)
        //        headerModel.Render(writer, document, ctx, parameters);
        //}

        ///// <summary>
        ///// Event raised just before the document.close()
        ///// </summary>
        ///// <param name="writer"></param>
        ///// <param name="document"></param>
        //public override void OnCloseDocument(PdfWriter writer, Document document)
        //{
        //    footerModel.OnCloseDocumentRender(writer, document);
        //}
    }
}
