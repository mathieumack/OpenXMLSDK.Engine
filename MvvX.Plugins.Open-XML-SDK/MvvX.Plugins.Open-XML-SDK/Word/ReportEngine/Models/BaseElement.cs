using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models
{
    public class BaseElement
    {
        /// <summary>
        /// Is the element visible
        /// </summary>
        public bool Show { get; set; } = true;

        /// <summary>
        /// Key from context used to determine element visibility
        /// </summary>
        public string ShowKey { get; set; }

        /// <summary>
        /// Childrens
        /// </summary>
        public IList<BaseElement> ChildElements { get; set; } = new List<BaseElement>();

        /// <summary>
        /// Font
        /// </summary>
        public string FontName { get; set; }

        /// <summary>
        /// Font size
        /// </summary>
        public string FontSize { get; set; }

        /// <summary>
        /// Font color
        /// </summary>
        public string FontColor { get; set; }
    }
}
