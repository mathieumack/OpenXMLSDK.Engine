using ReportEngine.Core.Template.Styles;

namespace OpenXMLSDK.Engine.Platform.Word.Extensions
{
    /// <summary>
    /// extension class for the stylevalues enumeration
    /// </summary>
    public static class StyleValuesExtensions
    {
        /// <summary>
        /// convert the StyleValue enum.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static DocumentFormat.OpenXml.Wordprocessing.StyleValues ToOOxml(this StyleValues value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.StyleValues)(int)value;
        }
    }
}
