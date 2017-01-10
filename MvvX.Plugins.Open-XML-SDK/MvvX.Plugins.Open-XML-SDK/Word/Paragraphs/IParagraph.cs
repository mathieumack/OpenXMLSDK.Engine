namespace MvvX.Plugins.OpenXMLSDK.Word.Paragraphs
{
    public interface IParagraph : IOpenXmlElement
    {
        /// <summary>
        /// Properties of the paragraph item
        /// </summary>
        IParagraphProperties ParagraphProperties { get; }
    }
}
