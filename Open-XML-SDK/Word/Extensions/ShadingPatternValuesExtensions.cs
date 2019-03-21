namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Extensions
{
    public static class ShadingPatternValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.ShadingPatternValues? ToOOxml(this MvvX.Plugins.OpenXMLSDK.Word.ShadingPatternValues? value)
        {
            if (value.HasValue)
                return (DocumentFormat.OpenXml.Wordprocessing.ShadingPatternValues)(int)value;
            else
                return null;
        }
    }
}
