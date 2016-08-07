namespace MvvX.Open_XML_SDK.Core.Word.Tables
{
    public interface ITableCellProperties : IOpenXmlElement
    {
        IGridSpan GridSpan { get; }

        Core.Word.OnOffOnlyValues? NoWrap { get; set; }
    }
}
