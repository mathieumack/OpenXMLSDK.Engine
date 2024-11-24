using OpenXMLSDK.Engine.Word;

namespace OpenXMLSDK.Engine.Platform.Word.Extensions
{
    public static class TextDirectionValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.TextDirectionValues? ToOOxml(this TextDirectionValues? value)
        {
            if (value.HasValue)
                return new DocumentFormat.OpenXml.Wordprocessing.TextDirectionValues(value.ToString().ToLower());
            else
                return null;
        }
    }
}
