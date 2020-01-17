using System;
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
            if (value.Count(f => f == '{') != value.Count(f => f == '}') || value.Count(f => f == '{') != substitutionTexts.Count)
                throw new ArgumentOutOfRangeException();

            IList<string> substitutionTextStrings = new List<string>();

            foreach (BaseModel baseModel in substitutionTexts)
            {
                var baseModelValue = string.Empty;
                if (baseModel is Base64ContentModel)
                    baseModelValue = (baseModel as Base64ContentModel).Base64Content;
                else if (baseModel is BooleanModel)
                    baseModelValue = (baseModel as BooleanModel).Value.ToString();
                else if (baseModel is ByteContentModel)
                    baseModelValue = (baseModel as ByteContentModel).Content.ToString();
                else if (baseModel is DateTimeModel)
                    baseModelValue = (baseModel as DateTimeModel).Value.ToString();
                else if (baseModel is DoubleModel)
                    baseModelValue = (baseModel as DoubleModel).Value.ToString();
                else if (baseModel is StringModel)
                    baseModelValue = (baseModel as StringModel).Value;

                substitutionTextStrings.Add(baseModelValue);
            }

            base.Value = string.Format(value, substitutionTextStrings.Select(x => x).ToArray());
        }

        #endregion
    }
}
