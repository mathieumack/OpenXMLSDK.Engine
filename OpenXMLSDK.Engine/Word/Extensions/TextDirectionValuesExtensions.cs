using ReportEngine.Core.Template;

namespace OpenXMLSDK.Engine.Platform.Word.Extensions
{
    public static class TextDirectionValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.TextDirectionValues? ToOOxml(this TextDirectionValues? value)
        {
            if (value.HasValue)
                return (DocumentFormat.OpenXml.Wordprocessing.TextDirectionValues)(int)value;
            else
                return null;
        }
    }
}
