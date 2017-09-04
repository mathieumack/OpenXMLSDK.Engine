namespace MvvX.Plugins.OpenXMLSDK.Word.Paragraphs
{
    public interface IParagraphBorders
    {
        IBorderType TopBorder { get; }

        IBorderType RightBorder { get; }

        IBorderType BottomBorder { get; }

        IBorderType LeftBorder { get; }

        IBorderType InsideHorizontalBorder { get; }

        IBorderType InsideVerticalBorder { get; }
    }
}
