using System;
using System.Collections.Generic;

namespace ReportEngine.Core.Template
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
        /// Table of contents title
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Title style id
        /// </summary>
        public string TitleStyleId { get; set; }

        /// <summary>
        /// Style Id of the table of contents 
        /// </summary>
        public List<string> ToCStylesId { get; set; }

        /// <summary>
        /// Tab stop leader character value
        /// </summary>
        public TabStopLeaderCharValues LeaderCharValue { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public TableOfContents() : base(typeof(TableOfContents).Name)
        {
            StylesAndLevels = new List<Tuple<string, string>>();
            ToCStylesId = new List<string>();
        }
    }
}
