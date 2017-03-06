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
    }
}
