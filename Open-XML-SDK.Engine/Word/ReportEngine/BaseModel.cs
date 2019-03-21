namespace OpenXMLSDK.Engine.Word.ReportEngine
{
    /// <summary>
    /// base class for context element
    /// </summary>
    public abstract class BaseModel
    {
        /// <summary>
        /// Type name of class : used for json serialization
        /// </summary>
        public string TypeName { get; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="type"></param>
        public BaseModel(string type)
        {
            TypeName = type;
        }
    }
}
