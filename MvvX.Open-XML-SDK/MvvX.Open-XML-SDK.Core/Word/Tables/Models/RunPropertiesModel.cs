using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MvvX.Open_XML_SDK.Core.Word.Tables.Models
{
    public class RunPropertiesModel
    {
        public bool? Bold { get; set; }

        public bool? Italic { get; set; }

        public string Color { get; set; }

        public string FontFamily { get; set; }

        public string FontSize { get; set; }

        public RunPropertiesModel()
        {

        }
    }
}
