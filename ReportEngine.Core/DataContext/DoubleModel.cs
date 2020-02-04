using System;

namespace ReportEngine.Core.DataContext
{
    public class DoubleModel : BaseModel
    {
        /// <summary>
        /// Value (0 by default)
        /// </summary>
        public double Value { get; set; }

        /// <summary>
        /// Used to define the final rendering string You can set precision of other :
        /// Ex : '{0:G2} kV' for the value 3.230 will generate '3.23 kV' string
        /// </summary>
        public string RenderPattern { get; set; }

        #region Constructors

        /// <summary>
        /// Default Constructor
        /// </summary>
        public DoubleModel()
            : this(0, null)
        { }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="value"></param>
        public DoubleModel(double value, string renderPattern)
            : base(typeof(DoubleModel).Name)
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
                return string.Format(formatProvider, RenderPattern, Value);
            else
                return Value.ToString(formatProvider);
        }

        #endregion
    }
}
