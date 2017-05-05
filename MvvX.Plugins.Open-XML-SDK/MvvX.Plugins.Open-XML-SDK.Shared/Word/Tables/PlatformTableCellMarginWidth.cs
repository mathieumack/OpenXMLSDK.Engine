using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.Tables;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Tables
{
    public class PlatformTableCellMarginWidth<T> : PlatformOpenXmlElement, ITableWidth
        where T : TableWidthType, new()
    {
        private readonly T xmlElement;

        public PlatformTableCellMarginWidth()
            : this(new T())
        {
        }

        public PlatformTableCellMarginWidth(T xmlElement)
            : base(xmlElement)
        {
            this.xmlElement = xmlElement;
        }

        public static PlatformTableCellMarginWidth<T> New(TableCellMargin tableCellMargin)
        {
            T xmlElement = null;
            if (tableCellMargin.Descendants<T>().Any())
                xmlElement = tableCellMargin.Descendants<T>().First();
            else
            {
                xmlElement = new T();
                tableCellMargin.Append(xmlElement);
            }
            return new PlatformTableCellMarginWidth<T>(xmlElement);
        }

        public OpenXMLSDK.Word.Tables.TableWidthUnitValues? Type
        {
            get
            {
                if (xmlElement.Type.HasValue)
                    return (OpenXMLSDK.Word.Tables.TableWidthUnitValues)(int)xmlElement.Type.Value;
                else
                    return null;
            }

            set
            {
                if (!value.HasValue)
                    xmlElement.Type = null;
                else
                    xmlElement.Type = (DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues)(int)value.Value;
            }
        }

        public string Width
        {
            get
            {
                return xmlElement.Width;
            }

            set
            {
                xmlElement.Width = value;
            }
        }
    }
}
