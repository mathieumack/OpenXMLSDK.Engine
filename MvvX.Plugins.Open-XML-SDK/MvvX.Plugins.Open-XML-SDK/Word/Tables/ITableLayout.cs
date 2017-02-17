namespace MvvX.Plugins.OpenXMLSDK.Word.Tables
{
    public interface ITableLayout : IOpenXmlElement
    {
        TableLayoutValues? Type { get; set; }
    }
}
