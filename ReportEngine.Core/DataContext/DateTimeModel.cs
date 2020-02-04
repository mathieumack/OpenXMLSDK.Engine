using System;

namespace ReportEngine.Core.DataContext
{
    public class DateTimeModel : BaseModel
    {
        /// <summary>
        /// Value (DateTime.MinValue by default)
        /// </summary>
        public DateTime Value { get; set; }

        /// <summary>
        /// Used to define the final rendering string You can set precision of other :
        /// More infos : https://msdn.microsoft.com/en-us/library/zdtaw1bw(v=vs.110).aspx
        /// </summary>
        public string RenderPattern { get; set; }

        #region Constructors

        /// <summary>
        /// Default Constructor
        /// </summary>
        public DateTimeModel()
            : this(DateTime.MinValue, null)
        { }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="value"></param>
        public DateTimeModel(DateTime value, string renderPattern)
            : base(typeof(DateTimeModel).Name)
        {
            Value = value;
            RenderPattern = renderPattern;
        }

        #endregion

        #region Rendering

        /// <summary>
        /// Create the generated string
        /// </summary>
        /// <returns></returns>
        public string Render()
        {
            return string.Format(RenderPattern, Value);
        }

        /// <summary>
        /// Create the generated string
        /// </summary>
        /// <param name="formatProvider">Format provider to be used for rendering</param>
        /// <returns></returns>
        public string Render(IFormatProvider formatProvider)
        {
            if (!string.IsNullOrWhiteSpace(RenderPattern))
                return Value.ToString(RenderPattern, formatProvider);
            else
                return Value.ToString(formatProvider);
        }

        #endregion
    }
}
