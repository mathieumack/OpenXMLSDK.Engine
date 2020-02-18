using ReportEngine.Core.Template.Charts;

namespace OpenXMLSDK.Engine.Word.Extensions
{
    public static class DataLabelPositionValuesExtensions
    {
        /// <summary>
        /// Convert the <see cref="DataLabelPositionValues"/> enumeration to the OpenXml <see cref="DocumentFormat.OpenXml.Drawing.Charts.DataLabelPositionValues" /> enumeration value
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static DocumentFormat.OpenXml.Drawing.Charts.DataLabelPositionValues ToOOxml(this DataLabelPositionValues value)
        {
            return (DocumentFormat.OpenXml.Drawing.Charts.DataLabelPositionValues)(int)value;
        }
    }
}
