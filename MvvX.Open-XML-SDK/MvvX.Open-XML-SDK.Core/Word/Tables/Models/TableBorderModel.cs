using System;
using MvvX.Open_XML_SDK.Core.Word.Bases;

namespace MvvX.Open_XML_SDK.Core.Word.Tables.Models
{
    public class TableBorderModel
    {
        public IOpenXmlElement[] childElements;
        public BorderValue BorderValue { get; set; }

        public UInt32 Size { get; set; }

        public string Color { get; set; }

        public TableBorderModel()
        {
            Size = 1;
            Color = "000000";
            BorderValue = BorderValue.Single;
        }

    }
}
