using OpenXMLSDK.Engine.Word;

namespace OpenXMLSDK.Engine.Platform.Word.Extensions
{
    public static class ShadingPatternValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.ShadingPatternValues? ToOOxml(this ShadingPatternValues? value)
        {
            if (value.HasValue)
                return new DocumentFormat.OpenXml.Wordprocessing.ShadingPatternValues(value.ToString().ToLower());
            else
                return null;
        }
    }
}
