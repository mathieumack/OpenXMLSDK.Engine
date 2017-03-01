using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.BatchModels
{
    public class BooleanModel : BaseModel
    {
        /// <summary>
        /// Value
        /// </summary>
        public bool Value { get; set; }

        #region Constructors
        public BooleanModel()
            : this(false)
        { }

        public BooleanModel(bool value)
            : base( typeof(BooleanModel).Name)
        {
            this.Value = value;
        }
        #endregion
    }
}
