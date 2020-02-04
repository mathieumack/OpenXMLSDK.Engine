using ReportEngine.Core.Template.Tables;

namespace OpenXMLSDK.Engine.Platform.Word.Extensions
{
    public static class TableVerticalAlignementValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.TableVerticalAlignmentValues ToOOxml(this TableVerticalAlignmentValues value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.TableVerticalAlignmentValues)(int)value;
        }
    }
}
