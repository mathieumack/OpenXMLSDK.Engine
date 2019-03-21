namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Extensions
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
        public static DocumentFormat.OpenXml.Wordprocessing.StyleValues ToOOxml(this MvvX.Plugins.OpenXMLSDK.Word.StyleValues value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.StyleValues)(int)value;
        }
    }
}
