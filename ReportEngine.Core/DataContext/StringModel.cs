namespace ReportEngine.Core.DataContext
{
    /// <summary>
    /// Model class for a string value in context
    /// </summary>
    public class StringModel : BaseModel
    {
        /// <summary>
        /// string value
        /// </summary>
        public string Value { get; set; }

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public StringModel()
            : this(null)
        { }

        /// <summary>
        /// constructor
        /// </summary>
        /// <param name="value"></param>
        public StringModel(string value)
            : base(typeof(StringModel).Name)
        {
            Value = value;
        }

        #endregion
    }
}
