using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models
{
    public class Document : BaseElement
    {
        /// <summary>
        /// List of pages of document
        /// </summary>
        public IList<Page> Pages { get; set; } = new List<Page>();
    }
}
