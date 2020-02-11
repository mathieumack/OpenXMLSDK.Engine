using System;
using System.Text.RegularExpressions;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template.Charts;
using ReportEngine.Core.Template.ExtendedModels;
using ReportEngine.Core.Template.Images;
using ReportEngine.Core.Template.Tables;
using ReportEngine.Core.Template.Text;

namespace ReportEngine.Core.Template.Extensions
{
    public static class ContextModelExtensions
    {
        public static string ReplaceText(this ContextModel context, string value, IFormatProvider formatProvider)
        {
            var result = value;
            var pattern = @"(#[a-zA-Z\.\\_\{\}0-9]+#)";
            if (Regex.IsMatch(value, pattern))
            {
                foreach (Match match in Regex.Matches(value, pattern))
                {
                    var key = match.Value;
                    if (context.ExistItem<StringModel>(key))
                        result = result.Replace(key, context.GetItem<StringModel>(key).Value);
                    else if (context.ExistItem<DoubleModel>(key))
                        result = result.Replace(key, context.GetItem<DoubleModel>(key).Render(formatProvider));
                    else if (context.ExistItem<DateTimeModel>(key))
                        result = result.Replace(key, context.GetItem<DateTimeModel>(key).Render(formatProvider));
                    else
                        result = result.Replace(key, string.Empty);
                }
            }
            return result;
        }

        /// <summary>
        /// Replace text and visibility of element
        /// </summary>
        /// <param name="element"></param>
        /// <param name="formatProvider"></param>
        public static void ReplaceItem(this ContextModel context, BarModel element, IFormatProvider formatProvider)
        {
            context.SetVisibilityFromContext(element);

            if (!string.IsNullOrEmpty(element.Title))
                element.Title = context.ReplaceText(element.Title, formatProvider);
            if (!string.IsNullOrEmpty(element.DataLabelColor))
                element.DataLabelColor = context.ReplaceText(element.DataLabelColor, formatProvider);
            if (!string.IsNullOrEmpty(element.FontName))
                element.FontName = context.ReplaceText(element.FontName, formatProvider);
            if (!string.IsNullOrEmpty(element.FontColor))
                element.FontColor = context.ReplaceText(element.FontColor, formatProvider);
            if (!string.IsNullOrEmpty(element.MajorGridlinesColor))
                element.MajorGridlinesColor = context.ReplaceText(element.MajorGridlinesColor, formatProvider);
        }

        /// <summary>
        /// Replace text and visibility of element
        /// </summary>
        /// <param name="element"></param>
        /// <param name="formatProvider"></param>
        public static void ReplaceItem(this ContextModel context, BaseElement element, IFormatProvider formatProvider)
        {
            context.SetVisibilityFromContext(element);
        }

        /// <summary>
        /// Replace an image by it context value
        /// </summary>
        /// <param name="element">template image object</param>
        public static void ReplaceItem(this ContextModel context, Image element)
        {
            if (!string.IsNullOrWhiteSpace(element.ContextKey))
            {
                // On va tenter de charger l'objet depuis sa base64 :
                if (context.ExistItem<Base64ContentModel>(element.ContextKey))
                {
                    var item = context.GetItem<Base64ContentModel>(element.ContextKey);
                    if (!string.IsNullOrWhiteSpace(item.Base64Content))
                        element.Content = Convert.FromBase64String(item.Base64Content);
                }
                else if (context.ExistItem<ByteContentModel>(element.ContextKey))
                {
                    var item = context.GetItem<ByteContentModel>(element.ContextKey);
                    element.Content = item.Content;
                }
                else if (context.ExistItem<FileLinkModel>(element.ContextKey))
                {
                    var item = context.GetItem<FileLinkModel>(element.ContextKey);
                    element.Path = item.Value;
                }
            }
            element.ContextKey = string.Empty;

            context.SetVisibilityFromContext(element);
        }

        /// <summary>
        /// Replace text, FontColor and visibility of element
        /// </summary>
        /// <param name="element"></param>
        /// <param name="formatProvider"></param>
        public static void ReplaceItem(this ContextModel context, Label element, IFormatProvider formatProvider)
        {
            if (!string.IsNullOrEmpty(element.Text))
                element.Text = context.ReplaceText(element.Text, formatProvider);
            if (!string.IsNullOrEmpty(element.FontColor))
                element.FontColor = context.ReplaceText(element.FontColor, formatProvider);
            if (!string.IsNullOrEmpty(element.Shading))
                element.Shading = context.ReplaceText(element.Shading, formatProvider);
            if (!string.IsNullOrEmpty(element.BoldKey) && context.ExistItem<BooleanModel>(element.BoldKey))
            {
                var item = context.GetItem<BooleanModel>(element.BoldKey);
                element.Bold = item.Value;
            }
            context.SetVisibilityFromContext(element);
        }

        /// <summary>
        /// Replace text, FontColor and visibility of element
        /// </summary>
        /// <param name="element"></param>
        /// <param name="formatProvider"></param>
        public static void ReplaceItem(this ContextModel context, HtmlContent element, IFormatProvider formatProvider)
        {
            if (!string.IsNullOrEmpty(element.Text))
                element.Text = context.ReplaceText(element.Text, formatProvider);

            context.SetVisibilityFromContext(element);
        }

        /// <summary>
        /// Replace text, FontColor and visibility of element
        /// </summary>
        /// <param name="element"></param>
        /// <param name="formatProvider"></param>
        public static void ReplaceItem(this ContextModel context, BookmarkStart element, IFormatProvider formatProvider)
        {
            if (!string.IsNullOrEmpty(element.Id))
                element.Id = context.ReplaceText(element.Id, formatProvider);
            if (!string.IsNullOrEmpty(element.Name))
                element.Name = context.ReplaceText(element.Name, formatProvider);
            //if (!string.IsNullOrEmpty(element.BoldKey) && context.ExistItem<BooleanModel>(element.BoldKey))
            //{
            //    var item = context.GetItem<BooleanModel>(element.BoldKey);
            //    element.Bold = item.Value;
            //}
            context.SetVisibilityFromContext(element);
        }

        /// <summary>
        /// Replace text, FontColor and visibility of element
        /// </summary>
        /// <param name="element"></param>
        /// <param name="formatProvider"></param>
        public static void ReplaceItem(this ContextModel context, BookmarkEnd element, IFormatProvider formatProvider)
        {
            if (!string.IsNullOrEmpty(element.Id))
                element.Id = context.ReplaceText(element.Id, formatProvider);
            //if (!string.IsNullOrEmpty(element.BoldKey) && context.ExistItem<BooleanModel>(element.BoldKey))
            //{
            //    var item = context.GetItem<BooleanModel>(element.BoldKey);
            //    element.Bold = item.Value;
            //}
            context.SetVisibilityFromContext(element);
        }

        /// <summary>
        /// Replace anchor and doclocation and visibility of element
        /// </summary>
        /// <param name="element"></param>
        /// <param name="formatProvider"></param>
        public static void ReplaceItem(this ContextModel context, Hyperlink element, IFormatProvider formatProvider)
        {
            if (!string.IsNullOrEmpty(element.Anchor))
                element.Anchor = context.ReplaceText(element.Anchor, formatProvider);

            if (!string.IsNullOrEmpty(element.WebSiteUri))
                element.WebSiteUri = context.ReplaceText(element.WebSiteUri, formatProvider);

            //if (!string.IsNullOrEmpty(element.BoldKey) && context.ExistItem<BooleanModel>(element.BoldKey))
            //{
            //    var item = context.GetItem<BooleanModel>(element.BoldKey);
            //    element.Bold = item.Value;
            //}
            context.SetVisibilityFromContext(element);
        }

        /// <summary>
        /// Replace text, FontColor and visibility of element
        /// </summary>
        /// <param name="element"></param>
        /// <param name="formatProvider"></param>
        public static void ReplaceItem(this ContextModel context, Paragraph element, IFormatProvider formatProvider)
        {
            if (!string.IsNullOrEmpty(element.ParagraphStyleId))
                element.ParagraphStyleId = context.ReplaceText(element.ParagraphStyleId, formatProvider);
            if (!string.IsNullOrEmpty(element.FontColor))
                element.FontColor = context.ReplaceText(element.FontColor, formatProvider);
            if (!string.IsNullOrEmpty(element.Shading))
                element.Shading = context.ReplaceText(element.Shading, formatProvider);
            if (element.Borders != null)
                context.ReplaceBorders(element.Borders, formatProvider);
            //if (!string.IsNullOrEmpty(element.BoldKey) && context.ExistItem<BooleanModel>(element.BoldKey))
            //{
            //    var item = context.GetItem<BooleanModel>(element.BoldKey);
            //    element.Bold = item.Value;
            //}
            if (!string.IsNullOrEmpty(element.PageBreakBeforeKey) && context.ExistItem<BooleanModel>(element.PageBreakBeforeKey))
            {
                var item = context.GetItem<BooleanModel>(element.PageBreakBeforeKey);
                element.PageBreakBefore = item.Value;
            }
            context.SetVisibilityFromContext(element);
        }

        /// <summary>
        /// Replace text, FontColor and visibility of element
        /// </summary>
        /// <param name="element"></param>
        /// <param name="formatProvider"></param>
        public static void ReplaceItem(this ContextModel context, Cell element, IFormatProvider formatProvider)
        {
            if (!string.IsNullOrEmpty(element.FontColor))
                element.FontColor = context.ReplaceText(element.FontColor, formatProvider);
            if (!string.IsNullOrEmpty(element.Shading))
                element.Shading = context.ReplaceText(element.Shading, formatProvider);
            if (element.Borders != null)
                context.ReplaceBorders(element.Borders, formatProvider);
            if (!string.IsNullOrEmpty(element.FusionKey) && context.ExistItem<BooleanModel>(element.FusionKey))
            {
                var item = context.GetItem<BooleanModel>(element.FusionKey);
                element.Fusion = item.Value;
            }
            if (!string.IsNullOrEmpty(element.FusionChildKey) && context.ExistItem<BooleanModel>(element.FusionChildKey))
            {
                var item = context.GetItem<BooleanModel>(element.FusionChildKey);
                element.FusionChild = item.Value;
            }
            context.SetVisibilityFromContext(element);
        }

        /// <summary>
        /// Replace text, FontColor and visibility of element
        /// </summary>
        /// <param name="element"></param>
        /// <param name="formatProvider"></param>
        public static void ReplaceItem(this ContextModel context, Table element, IFormatProvider formatProvider)
        {
            if (!string.IsNullOrEmpty(element.FontColor))
                element.FontColor = context.ReplaceText(element.FontColor, formatProvider);
            if (!string.IsNullOrEmpty(element.Shading))
                element.Shading = context.ReplaceText(element.Shading, formatProvider);
            if (element.Borders != null)
                context.ReplaceBorders(element.Borders, formatProvider);
            context.SetVisibilityFromContext(element);
        }

        private static void SetVisibilityFromContext(this ContextModel context, BaseElement element)
        {
            if (!string.IsNullOrWhiteSpace(element.ShowKey) && context.ExistItem<BooleanModel>(element.ShowKey))
            {
                var item = context.GetItem<BooleanModel>(element.ShowKey);
                element.Show = item.Value;
            }
            element.ShowKey = string.Empty;
        }

        private static void ReplaceBorders(this ContextModel context, BorderModel borders, IFormatProvider formatProvider)
        {
            if (!string.IsNullOrEmpty(borders.BorderColor))
                borders.BorderColor = context.ReplaceText(borders.BorderColor, formatProvider);

            if (!string.IsNullOrEmpty(borders.BorderLeftColor))
                borders.BorderLeftColor = context.ReplaceText(borders.BorderLeftColor, formatProvider);
            if (!string.IsNullOrEmpty(borders.BorderTopColor))
                borders.BorderTopColor = context.ReplaceText(borders.BorderTopColor, formatProvider);
            if (!string.IsNullOrEmpty(borders.BorderRightColor))
                borders.BorderRightColor = context.ReplaceText(borders.BorderRightColor, formatProvider);
            if (!string.IsNullOrEmpty(borders.BorderBottomColor))
                borders.BorderBottomColor = context.ReplaceText(borders.BorderBottomColor, formatProvider);
        }
    }
}
