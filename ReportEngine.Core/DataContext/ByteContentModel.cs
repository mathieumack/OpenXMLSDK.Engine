namespace ReportEngine.Core.DataContext
{
    /// <summary>
    /// ByteConentModel
    /// </summary>
    public class ByteContentModel : BaseModel
    {
        /// <summary>
        /// Content
        /// </summary>
        public byte[] Content { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public ByteContentModel()
            : this(null)
        { }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="content">Content value</param>
        public ByteContentModel(byte[] content)
            : base(typeof(ByteContentModel).Name)
        {
            Content = content;
        }
    }
}

