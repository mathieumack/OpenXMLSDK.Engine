namespace MvvX.Open_XML_SDK.Shared.Word.Extensions
{
    public static class TableRowAlignmentValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.TableRowAlignmentValues ToOOxml(this MvvX.Open_XML_SDK.Core.Word.TableRowAlignmentValues value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.TableRowAlignmentValues)(int)value;
        }
    }
}
