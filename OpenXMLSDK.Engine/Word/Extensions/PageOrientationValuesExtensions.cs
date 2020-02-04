using ReportEngine.Core.Template;

namespace OpenXMLSDK.Engine.Platform.Word.Extensions
{
    /// <summary>
    /// Extension class for Page orientation values
    /// </summary>
    public static class PageOrientationValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues ToOOxml(this PageOrientationValues value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues)(int)value;
        }
    }
}
