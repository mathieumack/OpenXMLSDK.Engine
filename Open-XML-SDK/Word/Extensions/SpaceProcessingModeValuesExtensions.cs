
namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Extensions
{
    public static class SpaceProcessingModeValuesExtensions
    {
        public static DocumentFormat.OpenXml.EnumValue<DocumentFormat.OpenXml.SpaceProcessingModeValues> ToOOxml(this MvvX.Plugins.OpenXMLSDK.Word.SpaceProcessingModeValues value)
        {
            return (DocumentFormat.OpenXml.SpaceProcessingModeValues)(int)value;
        }

        public static DocumentFormat.OpenXml.EnumValue<DocumentFormat.OpenXml.SpaceProcessingModeValues> ToOOxml(this MvvX.Plugins.OpenXMLSDK.Word.SpaceProcessingModeValues? value)
        {
            if (value.HasValue)
                return (DocumentFormat.OpenXml.SpaceProcessingModeValues)(int)value.Value;
            else
                return null;
        }

        public static MvvX.Plugins.OpenXMLSDK.Word.SpaceProcessingModeValues? ToPlatform(this DocumentFormat.OpenXml.EnumValue<DocumentFormat.OpenXml.SpaceProcessingModeValues> value)
        {
            if (value.HasValue)
                return (MvvX.Plugins.OpenXMLSDK.Word.SpaceProcessingModeValues)(int)value.Value;
            else
                return null;
        }
    }
}
