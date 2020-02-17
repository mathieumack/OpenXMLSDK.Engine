using System;
using System.IO;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.Extensions;
using ReportEngine.Core.Template.Images;
using it = iTextSharp.text;
using itp = iTextSharp.text.pdf;

namespace Pdf.Engine.ReportEngine.Renders
{
    internal static class ImageExtensions
    {
        public static void Render(this Image element, 
                                        Document document,
                                        itp.PdfWriter writer,
                                        ContextModel context,
                                        EngineContext ctx,
                                        IFormatProvider formatProvider)
        {
            context.ReplaceItem(element);

            ctx.Parents.Add(element);

            var image = (it.Image)null;
            //if (Element.Content != null)
            //    image = Image.GetInstance(Element.Content);
            if (!string.IsNullOrWhiteSpace(element.Path) && File.Exists(element.Path))
                image = it.Image.GetInstance(File.ReadAllBytes(element.Path));
            else if (element.Content != null)
                image = it.Image.GetInstance(element.Content);

            if (image != null && element.Show)
            {
                // TODO : Manage size
                // On va maintenant gérer le poucentage :
                var width = element.Width.HasValue ? element.Width : element.MaxWidth;
                if (element.MaxWidth.HasValue && width.HasValue && element.MaxWidth.Value < width.Value)
                    width = element.MaxWidth;
                if (width.HasValue)
                {
                    // On va recalculer le pourcentage automatiquement :
                    image.ScaleAbsoluteWidth(width.Value / 2);
                }

                var height = element.Height.HasValue ? element.Height : element.MaxHeight;
                if (element.MaxHeight.HasValue && height.HasValue && element.MaxHeight.Value < height.Value)
                    height = element.MaxHeight;
                if (height.HasValue)
                {
                    // On va recalculer le pourcentage automatiquement :
                    image.ScaleAbsoluteHeight(height.Value / 2);
                }

                //image.ScaleAbsolute(width.Value / 2, height.Value / 2);

                // insertion dans le document :
                element.AddToParentContainer(ctx, image);
            }

            ctx.Parents.RemoveAt(ctx.Parents.Count - 1);
        }
    }
}
