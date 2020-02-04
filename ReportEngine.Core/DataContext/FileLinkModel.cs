namespace ReportEngine.Core.DataContext
{
    /// <summary>
    /// FileLinkMode
    /// Represent the path of a file
    /// </summary>
    public class FileLinkModel : BaseModel
    {
        /// <summary>
        /// File path
        /// </summary>
        public string Value { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public FileLinkModel()
            : this(null)
        { }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="value">Content value</param>
        public FileLinkModel(string value)
            : base(typeof(FileLinkModel).Name)
        {
            Value = value;
        }
    }
}