using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models.Attributes;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models
{
    public class Cell : BaseElement
    {
        /// <summary>
        /// Borders
        /// </summary>
        public BorderModel Borders { get; set; }

        /// <summary>
        /// Colspan
        /// </summary>
        public int ColSpan { get; set; }

        /// <summary>
        /// Row Span : Cell is merged vertically : it's the first merged cell
        /// </summary>
        public bool Fusion { get; set; }

        /// <summary>
        /// Cell is a hidden merged part of a rowspan
        /// </summary>
        public bool FusionChild { get; set; }
    }
}
