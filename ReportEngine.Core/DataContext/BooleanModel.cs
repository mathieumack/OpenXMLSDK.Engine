namespace ReportEngine.Core.DataContext
{
    /// <summary>
    /// Model for a boolean vlaue
    /// </summary>
    public class BooleanModel : BaseModel
    {
        /// <summary>
        /// Value
        /// </summary>
        public bool Value { get; set; }

        #region Constructors
        /// <summary>
        /// Default Constructor
        /// </summary>
        public BooleanModel()
            : this(false)
        { }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="value"></param>
        public BooleanModel(bool value)
            : base(typeof(BooleanModel).Name)
        {
            Value = value;
        }
        #endregion
    }
}
