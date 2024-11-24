using OpenXMLSDK.Engine.Word.Tables;

namespace OpenXMLSDK.Engine.Platform.Word.Extensions
{
    public static class TableWidthUnitValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues ToOOxml(this TableWidthUnitValues value)
        {
            return new DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues(value.ToString().ToLower());
        }
    }
}
