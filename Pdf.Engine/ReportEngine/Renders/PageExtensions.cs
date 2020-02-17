using it = iTextSharp.text;
using itp = iTextSharp.text.pdf;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using System;
using System.Linq;

namespace Pdf.Engine.ReportEngine.Renders
{
    public static class PageExtensions
    {
        internal static void Render(this Page page, 
                                        Document document,
                                        itp.PdfWriter writer,
                                        it.Document pdfDocument,
                                        ContextModel context,
                                        EngineContext ctx,
                                        IFormatProvider formatProvider)
        {
            if (!string.IsNullOrWhiteSpace(page.ShowKey) && context.ExistItem<BooleanModel>(page.ShowKey) && !context.GetItem<BooleanModel>(page.ShowKey).Value)
                return;

            ctx.Parents.Add(page);
            foreach (var child in page.ChildElements)
            {
                child.Render(document, writer, pdfDocument, context, ctx, formatProvider);
            }
            if (ctx.Parents.Any())
                ctx.Parents.RemoveAt(ctx.Parents.Count - 1);
        }
    }
}
