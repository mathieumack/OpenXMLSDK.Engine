namespace MvvX.Plugins.OpenXMLSDK.Word.Tables
{
    public interface ITableJustification : IOpenXmlElement
    {
        EnumValue<TableRowAlignmentValues> Val { get; set; }
    }
}
