namespace MvvX.Plugins.OpenXMLSDK.Word.Tables
{
    public interface ITableProperties : IOpenXmlElement
    {
        ITableStyle TableStyle { get; }

        IShading Shading { get; }

        ITableBorders TableBorders { get; }

        ITableWidth TableWidth { get; }

        Word.Tables.TableRowAlignmentValues? TableJustification { get; set; }

        ITableLayout Layout { get; }
    }
}
