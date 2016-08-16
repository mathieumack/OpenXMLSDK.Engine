namespace MvvX.Open_XML_SDK.Core.Word.Tables
{
    public interface ITableProperties : IOpenXmlElement
    {
        ITableStyle TableStyle { get; }

        IShading Shading { get; }

        ITableBorders TableBorders { get; }

        ITableWidth TableWidth { get; }

        Core.Word.Tables.TableRowAlignmentValues? TableJustification { get; set; }
    }
}
