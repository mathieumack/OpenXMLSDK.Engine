using System;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word;

namespace MvvX.Open_XML_SDK.Shared.Word
{
    public class PlatformBorder<T> : PlatformOpenXmlElement, IBorderType 
        where T : BorderType, new()
    {
        private readonly T borderType;

        public PlatformBorder(T borderType)
            : base(borderType)
        {
            this.borderType = borderType;
        }

        #region Static helpers methods

        public static PlatformBorder<T> New()
        {
            return new PlatformBorder<T>(new T());
        }

        public static PlatformBorder<T> New(T borderType)
        {
            if (borderType == null)
                borderType = new T();
            return new PlatformBorder<T>(borderType);
        }

        #endregion

        #region Interface

        public Core.Word.BorderValues? BorderValue
        {
            get
            {
                if (borderType.Val == null || !borderType.Val.HasValue)
                    return null;
                else
                    return (Core.Word.BorderValues)(int)borderType.Val.Value;
            }
            set
            {
                if (value == null)
                    borderType.Val = null;
                else
                    borderType.Val = (DocumentFormat.OpenXml.Wordprocessing.BorderValues)(int)value;
            }
        }

        public string Color
        {
            get
            {
                return borderType.Color;
            }

            set
            {
                borderType.Color = value;
            }
        }

        public int? Size
        {
            get
            {
                if (borderType.Size.HasValue)
                    return (int?)borderType.Size.Value;
                else
                    return null;
            }
            set
            {
                borderType.Size = Convert.ToUInt32(value);
            }
        }

        #endregion
    }
}
