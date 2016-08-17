namespace MvvX.Plugins.OpenXMLSDK.Word.Tables
{
    public interface ITableCellWidth : IOpenXmlElement
    {
        string Width { get; set; }

        TableWidthUnitValues? Type { get; set; }
    }
}
