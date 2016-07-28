using System.Linq;

namespace MvvX.Open_XML_SDK.Core.Word.Tables.Models
{
    public class TableBorderModel
    {
        public BorderValues BorderValue { get; set; }

        public UInt32Value Size { get; set; }

        public string Color { get; set; }

        public TableBorderModel()
        {
            Size = 1;
            Color = EOWordColors.BlackColor;
            BorderValue = BorderValues.Single;
        }
    }
}
