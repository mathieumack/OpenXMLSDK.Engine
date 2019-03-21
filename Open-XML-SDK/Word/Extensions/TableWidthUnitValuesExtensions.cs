namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Extensions
{
    public static class TableWidthUnitValuesExtensions
    {
        public static bool OOxmlEquals(this OpenXMLSDK.Word.Tables.TableWidthUnitValues value, DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues compareValue)
        {
            return ((DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues)(int)value).Equals(compareValue);
        }

        public static DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues ToOOxml(this OpenXMLSDK.Word.Tables.TableWidthUnitValues value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues)(int)value;
        }
    }
}
