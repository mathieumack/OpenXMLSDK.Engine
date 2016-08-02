using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word.Bases;

namespace MvvX.Open_XML_SDK.Shared.Word.Bases
{
    public class PlatformDecimalNumberType : PlatformOpenXmlElement
    {
        public DecimalNumberType decimalNumberType { get; set; }

        public PlatformDecimalNumberType(DecimalNumberType decimalNumberType)
            : base(decimalNumberType)
        {
            this.decimalNumberType = decimalNumberType;
        }
    }
}
