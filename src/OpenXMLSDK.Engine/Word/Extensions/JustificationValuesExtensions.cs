using OpenXMLSDK.Engine.Word;

namespace OpenXMLSDK.Engine.Platform.Word.Extensions
{
    public static class JustificationValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.JustificationValues ToOOxml(this JustificationValues value)
        {
            return (DocumentFormat.OpenXml.Wordprocessing.JustificationValues)(int)value;
        }
    }
}
