namespace MvvX.Plugins.Open_XML_SDK.Shared.Word.Extensions
{
    public static class ShadingPatternValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.ShadingPatternValues? ToOOxml(this MvvX.Plugins.Open_XML_SDK.Core.Word.ShadingPatternValues? value)
        {
            if (value.HasValue)
                return (DocumentFormat.OpenXml.Wordprocessing.ShadingPatternValues)(int)value;
            else
                return null;
        }
    }
}
