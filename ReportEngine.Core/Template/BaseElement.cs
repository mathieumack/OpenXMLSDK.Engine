using System.Collections.Generic;

namespace ReportEngine.Core.Template
{
    /// <summary>
    /// Base Element for word template
    /// </summary>
    public class BaseElement : BaseModel
    {
        /// <summary>
        /// Is the element visible
        /// </summary>
        public bool Show { get; set; } = true;

        /// <summary>
        /// Key from context used to determine element visibility
        /// </summary>
        public string ShowKey { get; set; }

        /// <summary>
        /// Childrens
        /// </summary>
        public List<BaseElement> ChildElements { get; set; } = new List<BaseElement>();

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="type"></param>
        public BaseElement(string type) : base(type)
        {
        }
    }
}
