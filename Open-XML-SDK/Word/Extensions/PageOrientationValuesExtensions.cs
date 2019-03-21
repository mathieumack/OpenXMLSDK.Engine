namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Extensions
{
    /// <summary>
    /// Extension class for Page orientation values
    /// </summary>
    public static class PageOrientationValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues ToOOxml(this MvvX.Plugins.OpenXMLSDK.Word.PageOrientationValues value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues)(int)value;
        }
    }
}
