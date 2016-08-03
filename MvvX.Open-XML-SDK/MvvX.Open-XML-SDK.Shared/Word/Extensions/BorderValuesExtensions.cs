namespace MvvX.Open_XML_SDK.Shared.Word.Extensions
{
    public static class BorderValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.BorderValues ToOOxml(this MvvX.Open_XML_SDK.Core.Word.BorderValues value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.BorderValues)(int)value;
        }
    }
}
