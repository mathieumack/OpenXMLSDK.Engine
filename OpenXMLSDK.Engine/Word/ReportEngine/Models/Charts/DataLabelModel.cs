using System;
using System.Collections.Generic;
using System.Text;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts
{
    public class DataLabelModel : BaseElement
    {
        /// <summary>
        /// Indicate if labels data will be showed
        /// </summary>
        public bool ShowDataLabel { get; set; } = true;

        /// <summary>
        /// Indicate if category name will be showed
        /// </summary>
        public bool ShowCatName { get; set; } = true;

        /// <summary>
        /// Indicate if percentage will be showed
        /// </summary>
        public bool ShowPercent { get; set; } = true;

        /// <summary>
        /// Ctor
        /// </summary>
        public DataLabelModel() : this(null)
        {
        }

        /// <summary>
        /// Ctor
        /// </summary>
        public DataLabelModel(string type) : base(type)
        {
        }
    }
}
