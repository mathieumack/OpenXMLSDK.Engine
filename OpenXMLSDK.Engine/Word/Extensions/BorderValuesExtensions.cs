using ReportEngine.Core.Template;

namespace OpenXMLSDK.Engine.Platform.Word.Extensions
{
    public static class BorderValuesExtensions
    {
        /// <summary>
        /// Convert the <see cref="BorderValues"/> enumeration to the OpenXml <see cref="DocumentFormat.OpenXml.Wordprocessing.BorderValues" /> enumeration value
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static DocumentFormat.OpenXml.Wordprocessing.BorderValues ToOOxml(this BorderValues value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.BorderValues)(int)value;
        }
    }
}
