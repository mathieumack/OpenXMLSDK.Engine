namespace ReportEngine.Core.DataContext
{
    /// <summary>
    /// Base64Content model
    /// </summary>
    public class Base64ContentModel : BaseModel
    {
        /// <summary>
        /// Content
        /// </summary>
        public string Base64Content { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public Base64ContentModel()
            : this(null)
        { }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="base64Content">Content value</param>
        public Base64ContentModel(string base64Content)
            : base(typeof(Base64ContentModel).Name)
        {
            Base64Content = base64Content;
        }
    }
}