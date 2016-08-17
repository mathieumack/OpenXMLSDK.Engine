namespace MvvX.Plugins.OpenXMLSDK.Word.Tables
{
    public interface ITable : IOpenXmlElement
    {
        ITableProperties Properties { get; }
    }
}
