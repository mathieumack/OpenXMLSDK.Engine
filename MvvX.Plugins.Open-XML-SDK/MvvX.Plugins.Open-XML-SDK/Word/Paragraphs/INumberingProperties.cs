using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MvvX.Plugins.OpenXMLSDK.Word.Paragraphs
{
    public interface INumberingProperties : IOpenXmlElement
    {
        int? NumberingLevelReference { get; set; }

        int? NumberingId { get; set; }
    }
}
