namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Extensions
{
    public static class TextDirectionValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.TextDirectionValues? ToOOxml(this MvvX.Plugins.OpenXMLSDK.Word.TextDirectionValues? value)
        {
            if (value.HasValue)
                return (DocumentFormat.OpenXml.Wordprocessing.TextDirectionValues)(int)value;
            else
                return null;
        }
    }
}
