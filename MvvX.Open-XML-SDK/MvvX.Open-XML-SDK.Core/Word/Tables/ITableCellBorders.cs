namespace MvvX.Open_XML_SDK.Core.Word.Tables
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
