namespace MvvX.Plugins.OpenXMLSDK.Word.Tables
{
    public interface ITableCellProperties : IOpenXmlElement
    {
        ITableCellBorders TableCellBorders { get; }

        IGridSpan GridSpan { get; }

        ITableCellWidth TableCellWidth { get; }

        IShading Shading { get; }

        Word.OnOffOnlyValues? NoWrap { get; set; }

        ITableCellMargin TableCellMargin { get; }
    }
}
