namespace MvvX.Plugins.OpenXMLSDK.Word.Tables
{
    public interface ITableWidth : IOpenXmlElement
    {
        string Width { get; set; }

        TableWidthUnitValues? Type { get; set; }
    }
}
