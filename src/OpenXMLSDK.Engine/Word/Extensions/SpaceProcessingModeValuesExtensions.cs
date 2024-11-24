
using OpenXMLSDK.Engine.Word;

namespace OpenXMLSDK.Engine.Platform.Word.Extensions
{
    public static class SpaceProcessingModeValuesExtensions
    {
        public static DocumentFormat.OpenXml.EnumValue<DocumentFormat.OpenXml.SpaceProcessingModeValues> ToOOxml(this SpaceProcessingModeValues value)
        {
            return new DocumentFormat.OpenXml.SpaceProcessingModeValues(value.ToString().ToLower());
        }

        public static DocumentFormat.OpenXml.EnumValue<DocumentFormat.OpenXml.SpaceProcessingModeValues> ToOOxml(this SpaceProcessingModeValues? value)
        {
            if (value.HasValue)
                return new DocumentFormat.OpenXml.SpaceProcessingModeValues(value.Value.ToString().ToLower());
            else
                return null;
        }

        public static SpaceProcessingModeValues? ToPlatform(this DocumentFormat.OpenXml.EnumValue<DocumentFormat.OpenXml.SpaceProcessingModeValues> value)
        {
            if (value.HasValue && value.Value == DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve)
                return SpaceProcessingModeValues.Preserve;
            else
                return SpaceProcessingModeValues.Default;
        }
    }
}
