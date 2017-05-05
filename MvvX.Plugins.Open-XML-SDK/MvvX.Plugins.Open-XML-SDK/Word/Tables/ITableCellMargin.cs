namespace MvvX.Plugins.OpenXMLSDK.Word.Tables
{
    public interface ITableCellMargin
    {
        ITableWidth BottomMargin { get; }

        ITableWidth LeftMargin { get; }

        ITableWidth RightMargin { get; }

        ITableWidth TopMargin { get; }
    }
}
