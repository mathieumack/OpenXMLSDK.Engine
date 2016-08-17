namespace MvvX.Plugins.OpenXMLSDK.Word
{
    public interface IRun : IOpenXmlElement
    {
        /// <summary>
        /// Properties of the run item
        /// </summary>
        IRunProperties Properties { get; }
    }
}
