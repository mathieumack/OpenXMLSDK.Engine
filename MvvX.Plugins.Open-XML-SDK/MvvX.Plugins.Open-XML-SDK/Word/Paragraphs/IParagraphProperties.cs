namespace MvvX.Plugins.OpenXMLSDK.Word.Paragraphs
{
    public interface IParagraphProperties : IOpenXmlElement
    {
        INumberingProperties NumberingProperties { get; }

        IParagraphStyleId ParagraphStyleId { get; }

        ISpacingBetweenLines SpacingBetweenLines { get; }
    }
}