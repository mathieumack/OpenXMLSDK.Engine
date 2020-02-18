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

            if (!string.IsNullOrWhiteSpace(element.Path) && File.Exists(element.Path))
                image = it.Image.GetInstance(File.ReadAllBytes(element.Path));
            else if (element.Content != null)
                image = it.Image.GetInstance(element.Content);

            if (image != null && element.Show)
            {
                // TODO : Manage size
                var finalWidth = GetLength(element.Width, element.MaxWidth, image.Width);
                var finalHeight = GetLength(element.Height, element.MaxHeight, image.Height);

                image.ScalePercent(Math.Min(finalWidth, finalHeight));

                // insertion dans le document :
                element.AddToParentContainer(ctx, image);
            }

            ctx.Parents.RemoveAt(ctx.Parents.Count - 1);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="definedValue"></param>
        /// <param name="maxValue"></param>
        /// <param name="currentValue"></param>
        /// <returns></returns>
        private static float GetLength(long? definedValue, long? maxValue, float currentValue)
        {
            if (definedValue.HasValue && maxValue.HasValue && definedValue.Value < maxValue.Value)
                return ((float)definedValue.Value / 2f) * 100f / currentValue;
            else if (definedValue.HasValue && maxValue.HasValue)
                return ((float)maxValue.Value / 2f) * 100f / currentValue;
            else if (definedValue.HasValue)
                return ((float)definedValue.Value / 2) * 100f / currentValue;
            else if (maxValue.HasValue)
                return ((float)maxValue.Value / 2f) * 100f / currentValue;
            return 100f;
        }
    }
}
