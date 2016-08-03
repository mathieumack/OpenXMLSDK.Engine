using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word.Paragraphs;
using MvvX.Open_XML_SDK.Core.Word.Tables;

namespace MvvX.Open_XML_SDK.Core.Word.Bases
{
    public abstract class PlatformDecimalNumberType : PlatformOpenXmlElement, IDecimalNumberType
    {
        private readonly DecimalNumberType openXmlElement;

        public PlatformDecimalNumberType(DecimalNumberType openXmlElement)
            : base(openXmlElement)
        {
            this.openXmlElement = openXmlElement;
        }

        public int Val
        {
            get
            {
                if (openXmlElement.Val.HasValue)
                    return openXmlElement.Val.Value;
                else
                    return 0;
            }
            set
            {
                openXmlElement.Val = value;
            }
        }
    }
}
