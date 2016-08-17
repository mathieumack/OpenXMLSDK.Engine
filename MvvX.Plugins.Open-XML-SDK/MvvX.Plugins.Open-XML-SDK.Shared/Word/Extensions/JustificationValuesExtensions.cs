namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Extensions
{
    public static class JustificationValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.JustificationValues ToOOxml(this MvvX.Plugins.OpenXMLSDK.Word.JustificationValues value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.JustificationValues)(int)value;
        }
    }
}
