using ReportEngine.Core.Template.Charts;

namespace OpenXMLSDK.Engine.Word.Extensions
{
    public static class DataLabelPositionValuesExtensions
    {
        public static DocumentFormat.OpenXml.Drawing.Charts.DataLabelPositionValues ToOOxml(this DataLabelPositionValues value)
        {
            return (DocumentFormat.OpenXml.Drawing.Charts.DataLabelPositionValues)(int)value;
        }
    }
}
