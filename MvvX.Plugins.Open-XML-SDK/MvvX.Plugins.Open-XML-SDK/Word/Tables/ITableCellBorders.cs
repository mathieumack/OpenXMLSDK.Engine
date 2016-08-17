namespace MvvX.Plugins.OpenXMLSDK.Word.Tables
{
    public interface ITableCellBorders
    {
        IBorderType TopBorder { get; }

        IBorderType RightBorder { get; }

        IBorderType BottomBorder { get; }

        IBorderType LeftBorder { get; }

        IBorderType TopLeftToBottomRightCellBorder { get; }

        IBorderType TopRightToBottomLeftCellBorder { get; }
    }
}
