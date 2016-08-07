namespace MvvX.Open_XML_SDK.Core.Word.Tables
{
    public interface ITableProperties : IOpenXmlElement
    {
        ITableBorders TableBorders { get; }

        Core.Word.Tables.TableRowAlignmentValues? TableJustification { get; set; }
    }
}
