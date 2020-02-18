using ReportEngine.Core.Template;

namespace OpenXMLSDK.Engine.Platform.Word.Extensions
{
    public static class JustificationValuesExtensions
    {
        /// <summary>
        /// Convert the <see cref="JustificationValues"/> enumeration to the OpenXml <see cref="DocumentFormat.OpenXml.Wordprocessing.JustificationValues" /> enumeration value
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static DocumentFormat.OpenXml.Wordprocessing.JustificationValues ToOOxml(this JustificationValues value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.JustificationValues)(int)value;
        }
    }
}
