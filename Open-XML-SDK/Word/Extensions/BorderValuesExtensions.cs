namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Extensions
{
    public static class BorderValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.BorderValues ToOOxml(this MvvX.Plugins.OpenXMLSDK.Word.BorderValues value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.BorderValues)(int)value;
        }
    }
}
