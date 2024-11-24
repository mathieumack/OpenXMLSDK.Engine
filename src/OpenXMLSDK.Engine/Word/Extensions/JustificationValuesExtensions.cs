using OpenXMLSDK.Engine.Word;

namespace OpenXMLSDK.Engine.Platform.Word.Extensions
{
    public static class JustificationValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.JustificationValues ToOOxml(this JustificationValues value)
        {
            return new DocumentFormat.OpenXml.Wordprocessing.JustificationValues(value.ToString().ToLower());
        }
    }
}
