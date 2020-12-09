using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using OpenXMLSDK.Engine.Word.ReportEngine.Models;
using OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts;
using OpenXMLSDK.Engine.Word.ReportEngine.Models.ExtendedModels;

namespace OpenXMLSDK.Engine.ReportEngine.DataContext
{
    /// <summary>
    /// Context 
    /// </summary>
    public class ContextModel
    {
        #region Fields

        private Dictionary<string, BaseModel> datas;

        #endregion

        #region Properties

        /// <summary>
        /// Data in context (property used for json serialization)
        /// </summary>
        public Dictionary<string, BaseModel> Data { get { return datas; } set { datas = value; } }

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        public ContextModel()
        {
            datas = new Dictionary<string, BaseModel>(StringComparer.OrdinalIgnoreCase);
        }

        #endregion

        #region Methods

        public void AddItem<T>(string key, T item) where T : BaseModel
        {
            if (item != null)
            {
                if (datas.ContainsKey(key))
                    datas[key] = item;
                else
                    datas.Add(key, item);
            }
        }

        public bool ExistItem<T>(string key) where T : BaseModel
        {
            return TryGetItem<T>(key, out _);
        }

        public T GetItem<T>(string key) where T : BaseModel
        {
            TryGetItem(key, out T item);
            return item;
        }

        public bool TryGetItem<T>(string key, out T item) where T : BaseModel
        {
            if (!string.IsNullOrWhiteSpace(key) && datas.TryGetValue(key, out BaseModel rawModel))
            {
                item = rawModel as T;
            }
            else
            {
                item = default;
            }
            return item != default;
        }

        public string ReplaceText(string value, IFormatProvider formatProvider)
        {
            var result = value;
            var pattern = @"(#[a-zA-Z\.\\_\{\}0-9]+#)";
            if (Regex.IsMatch(value, pattern))
            {
                foreach (Match match in Regex.Matches(value, pattern))
                {
                    var key = match.Value;
                    if (TryGetItem(key, out StringModel stringItem))
                    {
                        result = result.Replace(key, stringItem.Value);
                    }
                    else if (TryGetItem(key, out DoubleModel doubleItem))
                    {
                        result = result.Replace(key, doubleItem.Render(formatProvider));
                    }
                    else if (TryGetItem(key, out DateTimeModel dateItem))
                    {
                        result = result.Replace(key, dateItem.Render(formatProvider));
                    }
                    else if (TryGetItem(key, out SubstitutableStringModel subStringItem))
                    {
                        result = result.Replace(key, subStringItem.Render(formatProvider));
                    }
                    else
                    {
                        result = result.Replace(key, string.Empty);
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// Replace text and visibility of element
        /// </summary>
        /// <param name="element"></param>
        /// <param name="formatProvider"></param>
        public void ReplaceItem(BarModel element, IFormatProvider formatProvider)
        {
            SetVisibilityFromContext(element);

            if (!string.IsNullOrEmpty(element.Title))
                element.Title = ReplaceText(element.Title, formatProvider);
            if (!string.IsNullOrEmpty(element.DataLabelColor))
                element.DataLabelColor = ReplaceText(element.DataLabelColor, formatProvider);
            if (!string.IsNullOrEmpty(element.FontName))
                element.FontName = ReplaceText(element.FontName, formatProvider);
            if (!string.IsNullOrEmpty(element.FontColor))
                element.FontColor = ReplaceText(element.FontColor, formatProvider);
        }

        /// <summary>
        /// Replace text and visibility of element
        /// </summary>
        /// <param name="element"></param>
        /// <param name="formatProvider"></param>
        public void ReplaceItem(BaseElement element, IFormatProvider formatProvider)
        {
            SetVisibilityFromContext(element);
        }

        /// <summary>
        /// Replace an image by it context value
        /// </summary>
        /// <param name="element">template image object</param>
        public void ReplaceItem(Image element)
        {
            if (!string.IsNullOrWhiteSpace(element.ContextKey))
            {
                // On va tenter de charger l'objet depuis sa base64 :
                if (TryGetItem(element.ContextKey, out Base64ContentModel base64Item))
                {
                    if (!string.IsNullOrWhiteSpace(base64Item.Base64Content))
                    {
                        element.Content = Convert.FromBase64String(base64Item.Base64Content);
                    }
                }
                else if (TryGetItem(element.ContextKey, out ByteContentModel byteItem))
                {
                    element.Content = byteItem.Content;
                }
                else if (TryGetItem(element.ContextKey, out FileLinkModel fileLinkItem))
                {
                    element.Path = fileLinkItem.Value;
                }
            }
            element.ContextKey = string.Empty;

            SetVisibilityFromContext(element);
        }

        /// <summary>
        /// Replace text, FontColor and visibility of element
        /// </summary>
        /// <param name="element"></param>
        /// <param name="formatProvider"></param>
        public void ReplaceItem(Label element, IFormatProvider formatProvider)
        {
            if (!string.IsNullOrEmpty(element.Text))
                element.Text = ReplaceText(element.Text, formatProvider);
            if (!string.IsNullOrEmpty(element.FontColor))
                element.FontColor = ReplaceText(element.FontColor, formatProvider);
            if (!string.IsNullOrEmpty(element.Shading))
                element.Shading = ReplaceText(element.Shading, formatProvider);
            if (TryGetItem(element.BoldKey, out BooleanModel item))
            {
                element.Bold = item.Value;
            }
            SetVisibilityFromContext(element);
        }

        /// <summary>
        /// Replace text, FontColor and visibility of element
        /// </summary>
        /// <param name="element"></param>
        /// <param name="formatProvider"></param>
        public void ReplaceItem(HtmlContent element, IFormatProvider formatProvider)
        {
            if (!string.IsNullOrEmpty(element.Text))
                element.Text = ReplaceText(element.Text, formatProvider);

            SetVisibilityFromContext(element);
        }

        /// <summary>
        /// Replace text, FontColor and visibility of element
        /// </summary>
        /// <param name="element"></param>
        /// <param name="formatProvider"></param>
        public void ReplaceItem(BookmarkStart element, IFormatProvider formatProvider)
        {
            if (!string.IsNullOrEmpty(element.Id))
                element.Id = ReplaceText(element.Id, formatProvider);
            if (!string.IsNullOrEmpty(element.Name))
                element.Name = ReplaceText(element.Name, formatProvider);
            if (TryGetItem(element.BoldKey, out BooleanModel item))
            {
                element.Bold = item.Value;
            }
            SetVisibilityFromContext(element);
        }

        /// <summary>
        /// Replace text, FontColor and visibility of element
        /// </summary>
        /// <param name="element"></param>
        /// <param name="formatProvider"></param>
        public void ReplaceItem(BookmarkEnd element, IFormatProvider formatProvider)
        {
            if (!string.IsNullOrEmpty(element.Id))
                element.Id = ReplaceText(element.Id, formatProvider);
            if (TryGetItem(element.BoldKey, out BooleanModel item))
            {
                element.Bold = item.Value;
            }
            SetVisibilityFromContext(element);
        }

        /// <summary>
        /// Replace anchor and doclocation and visibility of element
        /// </summary>
        /// <param name="element"></param>
        /// <param name="formatProvider"></param>
        public void ReplaceItem(Hyperlink element, IFormatProvider formatProvider)
        {
            if (!string.IsNullOrEmpty(element.Anchor))
                element.Anchor = ReplaceText(element.Anchor, formatProvider);

            if (!string.IsNullOrEmpty(element.WebSiteUri))
                element.WebSiteUri = ReplaceText(element.WebSiteUri, formatProvider);

            if (TryGetItem(element.BoldKey, out BooleanModel item))
            {
                element.Bold = item.Value;
            }
            SetVisibilityFromContext(element);
        }

        /// <summary>
        /// Replace Instruction and visibility of element
        /// </summary>
        /// <param name="element"></param>
        /// <param name="formatProvider"></param>
        public void ReplaceItem(SimpleField element, IFormatProvider formatProvider)
        {
            if (!string.IsNullOrEmpty(element.Instruction))
            {
                element.Instruction = ReplaceText(element.Instruction, formatProvider);
            }

            SetVisibilityFromContext(element);
        }

        /// <summary>
        /// Replace text, FontColor and visibility of element
        /// </summary>
        /// <param name="element"></param>
        /// <param name="formatProvider"></param>
        public void ReplaceItem(Paragraph element, IFormatProvider formatProvider)
        {
            if (!string.IsNullOrEmpty(element.ParagraphStyleId))
                element.ParagraphStyleId = ReplaceText(element.ParagraphStyleId, formatProvider);
            if (!string.IsNullOrEmpty(element.FontColor))
                element.FontColor = ReplaceText(element.FontColor, formatProvider);
            if (!string.IsNullOrEmpty(element.Shading))
                element.Shading = ReplaceText(element.Shading, formatProvider);
            if (element.Borders != null)
                ReplaceBorders(element.Borders, formatProvider);
            if (TryGetItem(element.BoldKey, out BooleanModel item))
            {
                element.Bold = item.Value;
            }
            if (TryGetItem(element.PageBreakBeforeKey, out BooleanModel pageBreakItem))
            {
                element.PageBreakBefore = pageBreakItem.Value;
            }
            SetVisibilityFromContext(element);
        }

        /// <summary>
        /// Replace text, FontColor and visibility of element
        /// </summary>
        /// <param name="element"></param>
        /// <param name="formatProvider"></param>
        public void ReplaceItem(Cell element, IFormatProvider formatProvider)
        {
            if (!string.IsNullOrEmpty(element.FontColor))
                element.FontColor = ReplaceText(element.FontColor, formatProvider);
            if (!string.IsNullOrEmpty(element.Shading))
                element.Shading = ReplaceText(element.Shading, formatProvider);
            if (element.Borders != null)
                ReplaceBorders(element.Borders, formatProvider);
            if (TryGetItem(element.FusionKey, out BooleanModel fusionItem))
            {
                element.Fusion = fusionItem.Value;
            }
            if (TryGetItem(element.FusionChildKey, out BooleanModel fusionChildItem))
            {
                element.FusionChild = fusionChildItem.Value;
            }
            if (TryGetItem(element.ColSpanKey, out DoubleModel colSpanItem))
            {
                element.ColSpan = (int)colSpanItem.Value;
            }
            SetVisibilityFromContext(element);
        }

        /// <summary>
        /// Replace text, FontColor and visibility of element
        /// </summary>
        /// <param name="element"></param>
        /// <param name="formatProvider"></param>
        public void ReplaceItem(Table element, IFormatProvider formatProvider)
        {
            if (!string.IsNullOrEmpty(element.FontColor))
                element.FontColor = ReplaceText(element.FontColor, formatProvider);
            if (!string.IsNullOrEmpty(element.Shading))
                element.Shading = ReplaceText(element.Shading, formatProvider);
            if (element.Borders != null)
                ReplaceBorders(element.Borders, formatProvider);
            SetVisibilityFromContext(element);
        }

        private void SetVisibilityFromContext(BaseElement element)
        {
            if (TryGetItem(element.ShowKey, out BooleanModel item))
            {
                element.Show = item.Value;
            }
            element.ShowKey = string.Empty;
        }

        private void ReplaceBorders(BorderModel borders, IFormatProvider formatProvider)
        {
            if (!string.IsNullOrEmpty(borders.BorderColor))
                borders.BorderColor = ReplaceText(borders.BorderColor, formatProvider);

            if (!string.IsNullOrEmpty(borders.BorderLeftColor))
                borders.BorderLeftColor = ReplaceText(borders.BorderLeftColor, formatProvider);
            if (!string.IsNullOrEmpty(borders.BorderTopColor))
                borders.BorderTopColor = ReplaceText(borders.BorderTopColor, formatProvider);
            if (!string.IsNullOrEmpty(borders.BorderRightColor))
                borders.BorderRightColor = ReplaceText(borders.BorderRightColor, formatProvider);
            if (!string.IsNullOrEmpty(borders.BorderBottomColor))
                borders.BorderBottomColor = ReplaceText(borders.BorderBottomColor, formatProvider);
        }

        #endregion
    }
}
