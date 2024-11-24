using OpenXMLSDK.Engine.Word;

namespace OpenXMLSDK.Engine.Platform.Word.Extensions
{
    public static class BorderValuesExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.BorderValues ToOOxml(this BorderValues value)
        {
            return new DocumentFormat.OpenXml.Wordprocessing.BorderValues(value.ToString().ToLower());
        }
    }
}
