using ReportEngine.Core.Template.Tables;

namespace OpenXMLSDK.Engine.Platform.Word.Extensions
{
    public static class TableWidthUnitValuesExtensions
    {
        public static bool OOxmlEquals(this TableWidthUnitValues value, TableWidthUnitValues compareValue)
        {
            return ((DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues)(int)value).Equals(compareValue);
        }

        public static DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues ToOOxml(this TableWidthUnitValues value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues)(int)value;
        }
    }
}
