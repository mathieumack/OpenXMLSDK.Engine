using System;
using System.Collections.Generic;
using System.Linq;

namespace ReportEngine.Core.DataContext
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
        /// <list type="bullet">
        /// <item>
        /// <description>'{0:G2} miles' for the value 3.230 will generate '3.23 miles' string</description>
        /// </item>
        /// </list>
        /// </summary>
        public string RenderPattern { get; set; }

        #region Constructor

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
            if (DataSource is null || DataSource.Data is null || DataSource.Data.Values is null)
                return string.Empty;

            var renders = new List<string>();
            foreach (var baseModel in DataSource.Data.Values.Where(e => e != null))
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

            try
            {
                return string.Format(RenderPattern, renders.ToArray());
            }
            // When there are less parameters than expected in RenderPattern string 
            catch (FormatException)
            {
                // We return renderPattern not formatting.
                return RenderPattern;
            }
        }

        #endregion
    }
}
