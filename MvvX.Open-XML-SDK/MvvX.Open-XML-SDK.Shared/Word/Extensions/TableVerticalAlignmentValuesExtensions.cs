namespace MvvX.Open_XML_SDK.Shared.Word.Extensions
{
    public static class TableVerticalAlignmentValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.TableVerticalAlignmentValues ToOOxml(this MvvX.Open_XML_SDK.Core.Word.TableVerticalAlignmentValues value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.TableVerticalAlignmentValues)(int)value;
        }
    }
}
