using ReportEngine.Core.Template.Styles;

namespace OpenXMLSDK.Engine.Platform.Word.Extensions
{
    public static class ShadingPatternValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.ShadingPatternValues? ToOOxml(this ShadingPatternValues? value)
        {
            if (value.HasValue)
                return (DocumentFormat.OpenXml.Wordprocessing.ShadingPatternValues)(int)value;
            else
                return null;
        }
    }
}
