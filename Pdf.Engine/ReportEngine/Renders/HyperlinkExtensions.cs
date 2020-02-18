using System;
using System.Linq;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.Extensions;
using ReportEngine.Core.Template.Text;
using it = iTextSharp.text;
using itp = iTextSharp.text.pdf;

namespace Pdf.Engine.ReportEngine.Renders
{
    public static class HyperlinkExtensions
    {
        internal static it.Anchor Render(this Hyperlink label,
                                                Document document,
                                                itp.PdfWriter writer,
                                                ContextModel context,
                                                EngineContext ctx,
                                                IFormatProvider formatProvider)
        {
            context.ReplaceItem(label, formatProvider);
            if (!label.Show)
                return null;

            // Transform label Text before rendering :
            ApplyTransformOperations(label.Text);

            var anchor = new it.Anchor(label.Text.Text, label.Text.GetFont());
            anchor.Reference = label.WebSiteUri;

            return anchor;
        }

        /// <summary>
        /// Apply transforme operation on the label before rendering
        /// </summary>
        /// <param name="label"></param>
        private static void ApplyTransformOperations(Label label)
        {
            if (!string.IsNullOrWhiteSpace(label.Text) && label.TransformOperations != null)
            {
                foreach (var operation in label.TransformOperations.Where(e => e != null))
                {
                    switch (operation.TransformOperationType)
                    {
                        case LabelTransformOperationType.ToUpper:
                            label.Text = label.Text.ToUpper();
                            break;
                        case LabelTransformOperationType.ToLower:
                            label.Text = label.Text.ToLower();
                            break;
                        case LabelTransformOperationType.ToUpperInvariant:
                            label.Text = label.Text.ToUpperInvariant();
                            break;
                        case LabelTransformOperationType.ToLowerInvariant:
                            label.Text = label.Text.ToLowerInvariant();
                            break;
                        case LabelTransformOperationType.Trim:
                            label.Text = label.Text.Trim();
                            break;
                        case LabelTransformOperationType.TrimStart:
                            label.Text = label.Text.TrimStart();
                            break;
                        case LabelTransformOperationType.TrimEnd:
                            label.Text = label.Text.TrimEnd();
                            break;
                        default:
                            break;
                    }
                }
            }
        }
    }
}
