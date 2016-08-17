using System;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word;
using System.Linq;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word
{
    public class PlatformBorder<T> : PlatformOpenXmlElement, IBorderType 
        where T : BorderType, new()
    {
        private readonly T borderType;

        public PlatformBorder()
            : this(new T())
        {
        }

        public PlatformBorder(T borderType)
            : base(borderType)
        {
            this.borderType = borderType;
        }

        #region Static helpers methods

        public static PlatformBorder<T> New(TableCellBorders tableCellBorders)
        {
            T xmlElement = null;
            if (tableCellBorders.Descendants<T>().Any())
                xmlElement = tableCellBorders.Descendants<T>().First();
            else
            {
                xmlElement = new T();
                tableCellBorders.Append(xmlElement);
            }
            return new PlatformBorder<T>(xmlElement);
        }

        public static PlatformBorder<T> New(TableBorders tableBorders)
        {
            T xmlElement = null;
            if (tableBorders.Descendants<T>().Any())
                xmlElement = tableBorders.Descendants<T>().First();
            else
            {
                xmlElement = new T();
                tableBorders.Append(xmlElement);
            }
            return new PlatformBorder<T>(xmlElement);
        }

        #endregion

        #region Interface

        public OpenXMLSDK.Word.BorderValues? BorderValue
        {
            get
            {
                if (borderType.Val == null || !borderType.Val.HasValue)
                    return null;
                else
                    return (OpenXMLSDK.Word.BorderValues)(int)borderType.Val.Value;
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
