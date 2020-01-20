using System;
using System.Collections.Generic;
using System.Linq;


namespace OpenXMLSDK.Engine.ReportEngine.DataContext
{
    /// <summary>
    /// Model class for a substitutable string value in context
    /// </summary>
    public class SubstitutableStringModel : BaseModel
    {
        /// <summary>
        /// List of substitution texts
        /// </summary>
        public ContextModel DataSource { get; set; }

        /// <summary>
        /// Used to define the final rendering string You can set precision of other :
        /// Ex : '{0:G2} kV' for the value 3.230 will generate '3.23 kV' string
        /// </summary>
        public string RenderPattern { get; set; }

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public SubstitutableStringModel() : this(null, null)
        { }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="value"></param>
        /// <param name="substitutionTexts"></param>
        public SubstitutableStringModel(string renderPattern, ContextModel dataSource) : base(typeof(SubstitutableStringModel).Name)
        {
            DataSource = dataSource;
            this.RenderPattern = renderPattern;
        }

        #endregion

        #region Rendering

        /// <summary>
        /// Create the generated string
        /// </summary>
        /// <returns></returns>
        public string Render(IFormatProvider formatProvider)
        {
            var renders = new List<string>();

            foreach (var baseModel in DataSource.Data.Values)
            {
                var resultItem = "";
                if (baseModel is DoubleModel)
                    resultItem = (baseModel as DoubleModel).Render(formatProvider);
                else if (baseModel is DateTimeModel)
                    resultItem = (baseModel as DateTimeModel).Render(formatProvider);
                else if (baseModel is SubstitutableStringModel)
                    resultItem = (baseModel as SubstitutableStringModel).Render(formatProvider);
                else if (baseModel is StringModel)
                    resultItem = (baseModel as StringModel).Value;
                renders.Add(resultItem);
            }

            return string.Format(RenderPattern, renders.ToArray());
        }

        #endregion
    }
}
