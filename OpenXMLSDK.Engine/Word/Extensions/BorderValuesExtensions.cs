using ReportEngine.Core.Template;

namespace OpenXMLSDK.Engine.Platform.Word.Extensions
{
    public static class BorderValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.BorderValues ToOOxml(this BorderValues value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.BorderValues)(int)value;
        }
    }
}
