using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Models.Charts
{
    public class ChartCategory : BaseElement
    {
        /// <summary>
        /// Name of the category
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Category color
        /// </summary>
        public string Color { get; set; }

        public ChartCategory() 
            : base(typeof(ChartCategory).Name)
        {
        }
    }
}
