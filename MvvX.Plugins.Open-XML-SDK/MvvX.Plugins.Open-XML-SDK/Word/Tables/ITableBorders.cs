namespace MvvX.Plugins.OpenXMLSDK.Word.Tables
{
    public interface ITableBorders
    {
        IBorderType TopBorder { get; }

        IBorderType RightBorder { get; }

        IBorderType BottomBorder { get; }

        IBorderType LeftBorder { get; }

        IBorderType InsideHorizontalBorder { get; }

        IBorderType InsideVerticalBorder { get; }
    }
}
