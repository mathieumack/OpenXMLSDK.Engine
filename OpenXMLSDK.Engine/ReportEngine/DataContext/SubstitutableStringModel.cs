using System.Collections.Generic;
using System.Linq;


namespace OpenXMLSDK.Engine.ReportEngine.DataContext
{
    /// <summary>
    /// Model class for a substitutable string value in context
    /// </summary>
    public class SubstitutableStringModel : StringModel
    {
        /// <summary>
        /// List of substitution texts
        /// </summary>
        public List<BaseModel> SubstitutionTexts { get; set; }

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public SubstitutableStringModel() : this(null, new List<BaseModel>())
        { }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="value"></param>
        public SubstitutableStringModel(string value, List<BaseModel> substitutionTexts) : base(typeof(SubstitutableStringModel).Name)
        {
            base.Value = value;
        }

        #endregion
    }
}
