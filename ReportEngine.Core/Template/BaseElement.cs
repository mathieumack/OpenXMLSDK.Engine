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

        #region Fields use to configure an element and it's children

        /// <summary>
        /// Is text contained in current element bold
        /// </summary>
        public bool? Bold { get; set; }

        /// <summary>
        /// Is text contained in current element Italic
        /// </summary>
        public bool? Italic { get; set; }

        /// <summary>
        /// Font
        /// </summary>
        public string FontName { get; set; }

        /// <summary>
        /// Font size. The size is define in demi-point
        /// </summary>
        public int? FontSize { get; set; }

        /// <summary>
        /// Font color in hex value (RRGGBB format)
        /// </summary>
        public string FontColor { get; set; }

        /// <summary>
        /// Font encoding
        /// <para>Used only for PDF reporting. Default value : Cp1252</para>
        /// </summary>
        public string FontEncoding { get; set; }

        /// <summary>
        /// Shading color in hex value (RRGGBB format)
        /// </summary>
        public string Shading { get; set; }

        #endregion

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
