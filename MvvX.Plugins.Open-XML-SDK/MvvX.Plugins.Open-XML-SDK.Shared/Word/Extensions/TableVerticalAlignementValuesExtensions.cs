namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Extensions
{
    public static class TableVerticalAlignementValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.TableVerticalAlignmentValues ToOOxml(this MvvX.Plugins.OpenXMLSDK.Word.Tables.TableVerticalAlignmentValues value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.TableVerticalAlignmentValues)(int)value;
        }
    }
}
