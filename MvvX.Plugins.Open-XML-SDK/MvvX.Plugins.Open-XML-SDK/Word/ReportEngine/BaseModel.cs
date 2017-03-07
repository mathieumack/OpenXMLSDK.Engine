namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine
{
    /// <summary>
    /// base class for context element
    /// </summary>
    public abstract class BaseModel
    {
        public string Type { get; }

        public BaseModel(string type)
        {
            Type = type;
        }
    }
}
