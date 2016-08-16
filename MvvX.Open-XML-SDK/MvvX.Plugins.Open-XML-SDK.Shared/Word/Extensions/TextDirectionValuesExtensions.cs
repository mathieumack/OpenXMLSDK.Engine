namespace MvvX.Plugins.Open_XML_SDK.Shared.Word.Extensions
{
    public static class TextDirectionValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.TextDirectionValues? ToOOxml(this MvvX.Plugins.Open_XML_SDK.Core.Word.TextDirectionValues? value)
        {
            if (value.HasValue)
                return (DocumentFormat.OpenXml.Wordprocessing.TextDirectionValues)(int)value;
            else
                return null;
        }
    }
}
