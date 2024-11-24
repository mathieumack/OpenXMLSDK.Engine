using OpenXMLSDK.Engine.Word.Tables;

namespace OpenXMLSDK.Engine.Platform.Word.Extensions
{
    public static class TableWidthUnitValuesExtensions
    {
        public static bool OOxmlEquals(this TableWidthUnitValues value, TableWidthUnitValues compareValue)
        {
            return ((int)value).Equals((int)compareValue);
        }

        public static DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues ToOOxml(this TableWidthUnitValues value)
        {
            return new DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues(value.ToString().ToLower());
        }
    }
}
