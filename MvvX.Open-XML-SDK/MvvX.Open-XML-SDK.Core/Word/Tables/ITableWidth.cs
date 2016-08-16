namespace MvvX.Open_XML_SDK.Core.Word.Tables
{
    public interface ITableWidth : IOpenXmlElement
    {
        string Width { get; set; }

        TableWidthUnitValues? Type { get; set; }
    }
}
