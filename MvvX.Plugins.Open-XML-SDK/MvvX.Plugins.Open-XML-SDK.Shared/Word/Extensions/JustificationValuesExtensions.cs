namespace MvvX.Plugins.Open_XML_SDK.Shared.Word.Extensions
{
    public static class JustificationValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.JustificationValues ToOOxml(this MvvX.Plugins.Open_XML_SDK.Core.Word.JustificationValues value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.JustificationValues)(int)value;
        }
    }
}
