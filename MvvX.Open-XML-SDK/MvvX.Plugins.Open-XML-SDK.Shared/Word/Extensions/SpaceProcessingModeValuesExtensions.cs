using DocumentFormat.OpenXml;

namespace MvvX.Plugins.Open_XML_SDK.Shared.Word.Extensions
{
    public static class SpaceProcessingModeValuesExtensions
    {
        public static EnumValue<DocumentFormat.OpenXml.SpaceProcessingModeValues> ToOOxml(this Core.Word.SpaceProcessingModeValues value)
        {
            return (DocumentFormat.OpenXml.SpaceProcessingModeValues)(int)value;
        }

        public static EnumValue<DocumentFormat.OpenXml.SpaceProcessingModeValues> ToOOxml(this Core.Word.SpaceProcessingModeValues? value)
        {
            if (value.HasValue)
                return (DocumentFormat.OpenXml.SpaceProcessingModeValues)(int)value.Value;
            else
                return null;
        }

        public static Core.Word.SpaceProcessingModeValues? ToPlatform(this EnumValue<DocumentFormat.OpenXml.SpaceProcessingModeValues> value)
        {
            if (value.HasValue)
                return (Core.Word.SpaceProcessingModeValues)(int)value.Value;
            else
                return null;
        }
    }
}
