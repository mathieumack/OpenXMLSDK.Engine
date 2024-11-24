using OpenXMLSDK.Engine.Word.Tables;

namespace OpenXMLSDK.Engine.Platform.Word.Extensions
{
    public static class TableVerticalAlignementValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.TableVerticalAlignmentValues ToOOxml(this TableVerticalAlignmentValues value)
        {
            return new DocumentFormat.OpenXml.Wordprocessing.TableVerticalAlignmentValues(value.ToString().ToLower());
        }
    }
}
