using iTextSharp.text.pdf;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using System;

namespace Pdf.Engine.ReportEngine.Renders
{
    public static class PageExtensions
    {
        internal static void Render(this Page page, 
                                        Document document,
                                        PdfWriter writer,
                                        ContextModel context,
                                        EngineContext ctx,
                                        IFormatProvider formatProvider)
        {
            if (!string.IsNullOrWhiteSpace(page.ShowKey) && context.ExistItem<BooleanModel>(page.ShowKey) && !context.GetItem<BooleanModel>(page.ShowKey).Value)
                return;
            
            // add page content
            ((BaseElement)page).Render(document, writer, context, ctx, formatProvider);
        }
    }
}
