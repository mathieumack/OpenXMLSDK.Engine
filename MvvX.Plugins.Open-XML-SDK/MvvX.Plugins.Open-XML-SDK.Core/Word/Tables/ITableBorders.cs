namespace MvvX.Plugins.Open_XML_SDK.Core.Word.Tables
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
