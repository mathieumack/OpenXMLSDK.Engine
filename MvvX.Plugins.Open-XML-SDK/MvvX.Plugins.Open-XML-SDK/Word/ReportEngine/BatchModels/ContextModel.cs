using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels
{
    /// <summary>
    /// Context 
    /// </summary>
    public class ContextModel
    {
        #region Fields

        private  Dictionary<string, BaseModel> datas;
        #endregion

        #region Properties
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

        public T GetItem<T>(string key) where T : BaseModel
        {
            if (datas.ContainsKey(key))
                return (T)datas[key];
            else
                return default(T);
        }

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
            return !string.IsNullOrWhiteSpace(key) && datas.ContainsKey(key) && datas[key] is T;
        }

        public string ReplaceText(string value)
        {
            var result = value;
            var pattern = @"(#[a-zA-Z\.\\_\{\}0-9]+#)";
            if (Regex.IsMatch(value, pattern))
            {
                foreach (Match match in Regex.Matches(value, pattern))
                {
                    var key = match.Value;
                    if (ExistItem<StringModel>(key))
                        result = result.Replace(key, GetItem<StringModel>(key).Value);
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
        public void ReplaceItem(BaseElement element)
        {
            SetVisibilityFromContext(element);
        }

        /// <summary>
        /// Replace text and visibility of element
        /// </summary>
        /// <param name="element"></param>
        public void ReplaceItem(Label element)
        {
            element.Text = ReplaceText(element.Text);
            SetVisibilityFromContext(element);
        }

        private void SetVisibilityFromContext(BaseElement element)
        {
            if (!string.IsNullOrWhiteSpace(element.ShowKey))
            {
                if (ExistItem<BooleanModel>(element.ShowKey))
                {
                    var item = GetItem<BooleanModel>(element.ShowKey);
                    element.Show = item.Value;
                }
            }
            element.ShowKey = string.Empty;
        }
        #endregion
    }
}
