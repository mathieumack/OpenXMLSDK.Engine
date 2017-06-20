using System;
using System.Collections.Generic;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models
{
    /// <summary>
    /// Table of contents model
    /// </summary>
    public class TableOfContents : BaseElement
    {
        /// <summary>
        /// List of styles to add in the ToC and associated level
        /// </summary>
        public List<Tuple<string, string>> StylesAndLevels { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="type"></param>
        public TableOfContents() : base(typeof(TableOfContents).Name)
        {
            StylesAndLevels = new List<Tuple<string, string>>();
        }
    }
}
