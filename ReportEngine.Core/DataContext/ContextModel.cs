using System;
using System.Collections.Generic;

namespace ReportEngine.Core.DataContext
{
    /// <summary>
    /// Context 
    /// </summary>
    public class ContextModel
    {
        #region Properties

        /// <summary>
        /// Data in context (property used for json serialization)
        /// </summary>
        public Dictionary<string, BaseModel> Data { get; set; } = new Dictionary<string, BaseModel>(StringComparer.OrdinalIgnoreCase);

        #endregion

        #region Methods

        public T GetItem<T>(string key) where T : BaseModel
        {
            if (Data.ContainsKey(key))
                return (T)Data[key];
            else
                return default(T);
        }

        public void AddItem<T>(string key, T item) where T : BaseModel
        {
            if (item != null)
            {
                if (Data.ContainsKey(key))
                    Data[key] = item;
                else
                    Data.Add(key, item);
            }
        }

        public bool ExistItem<T>(string key) where T : BaseModel
        {
            return !string.IsNullOrWhiteSpace(key) && Data.ContainsKey(key) && Data[key] is T;
        }
        
        #endregion
    }
}
