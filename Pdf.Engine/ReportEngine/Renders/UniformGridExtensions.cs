using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Pdf.Engine.ReportEngine.Extensions;
using Pdf.Engine.ReportEngine.Helpers;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.Extensions;
using ReportEngine.Core.Template.Tables;
using it = iTextSharp.text;
using itp = iTextSharp.text.pdf;

namespace Pdf.Engine.ReportEngine.Renders
{
    internal static class UniformGridExtensions
    {
        /// <summary>
        /// Déclanche le rendu de l'élément
        /// </summary>
        /// <param name="table"></param>
        public static void Render(this UniformGrid uniformGrid,
                                        Document document,
                                        itp.PdfWriter writer,
                                        it.Document pdfDocument,
                                        ContextModel context,
                                        EngineContext ctx,
                                        IFormatProvider formatProvider)
        {
            context.ReplaceItem(uniformGrid, formatProvider);

            if (!string.IsNullOrEmpty(uniformGrid.DataSourceKey) && context.ExistItem<DataSourceModel>(uniformGrid.DataSourceKey))
            {
                
            }
        }
    }
}
