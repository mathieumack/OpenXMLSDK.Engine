﻿namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models
{
    public class ForEach : BaseElement
    {
        /// <summary>
        /// Describe the template of each final item
        /// </summary>
        public BaseElement ItemTemplate { get; set; }

        /// <summary>
        /// Items that will be repeated
        /// </summary>
        public string DataSourceKey { get; set; }

        public ForEach()
            : base(typeof(ForEach).Name)
        {
        }
    }
}