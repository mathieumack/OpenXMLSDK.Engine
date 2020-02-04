
using ReportEngine.Core.Template.Text;

namespace OpenXMLSDK.Engine.Platform.Word.Extensions
{
    public static class SpaceProcessingModeValuesExtensions
    {
        public static DocumentFormat.OpenXml.EnumValue<DocumentFormat.OpenXml.SpaceProcessingModeValues> ToOOxml(this SpaceProcessingModeValues value)
        {
            return (DocumentFormat.OpenXml.SpaceProcessingModeValues)(int)value;
        }

        public static DocumentFormat.OpenXml.EnumValue<DocumentFormat.OpenXml.SpaceProcessingModeValues> ToOOxml(this SpaceProcessingModeValues? value)
        {
            if (value.HasValue)
                return (DocumentFormat.OpenXml.SpaceProcessingModeValues)(int)value.Value;
            else
                return null;
        }

        public static SpaceProcessingModeValues? ToPlatform(this DocumentFormat.OpenXml.EnumValue<DocumentFormat.OpenXml.SpaceProcessingModeValues> value)
        {
            if (value.HasValue)
                return (SpaceProcessingModeValues)(int)value.Value;
            else
                return null;
        }
    }
}
