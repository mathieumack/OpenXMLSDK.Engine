using System;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word;
using MvvX.Open_XML_SDK.Core.Word.Bases;
using DocumentFormat.OpenXml;

namespace MvvX.Open_XML_SDK.Shared.Word
{
    public abstract class PlatformOnOffType : PlatformOpenXmlElement, IOnOffType
    {
        private readonly OnOffType onOffType;

        public PlatformOnOffType(OnOffType onOffType)
            : base(onOffType)
        {
            this.onOffType = onOffType;
        }

        #region Interface

        public bool Value
        {
            get
            {
                return OnOffValue.ToBoolean(onOffType.Val);
            }

            set
            {
                onOffType.Val = OnOffValue.FromBoolean(value);
            }
        }

        #endregion
    }
}
