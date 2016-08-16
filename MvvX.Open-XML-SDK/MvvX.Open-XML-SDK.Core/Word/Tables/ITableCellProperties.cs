namespace MvvX.Open_XML_SDK.Core.Word.Tables
{
    public interface ITableCellProperties : IOpenXmlElement
    {
        ITableCellBorders TableCellBorders { get; }
        
        IGridSpan GridSpan { get; }

        ITableCellWidth TableCellWidth { get; }

        IShading Shading { get; }

        Core.Word.OnOffOnlyValues? NoWrap { get; set; }
    }
}
