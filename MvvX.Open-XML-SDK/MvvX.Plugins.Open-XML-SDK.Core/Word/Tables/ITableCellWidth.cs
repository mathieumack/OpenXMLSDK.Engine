namespace MvvX.Plugins.Open_XML_SDK.Core.Word.Tables
{
    public interface ITableCellWidth : IOpenXmlElement
    {
        string Width { get; set; }

        TableWidthUnitValues? Type { get; set; }
    }
}
