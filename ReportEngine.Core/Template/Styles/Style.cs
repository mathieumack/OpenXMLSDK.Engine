namespace ReportEngine.Core.Template.Styles
{
    /// <summary>
    /// Model class for a style in a word document
    /// </summary>
    public class Style : BaseElement
    {
        /// <summary>
        /// Id of style
        /// </summary>
        public string StyleId { get; set; }

        /// <summary>
        /// Type of style
        /// </summary>
        public StyleValues Type { get; set; }

        /// <summary>
        /// Id of the based on style
        /// </summary>
        public string StyleBasedOn { get; set; }

        /// <summary>
        /// Custom style flag
        /// </summary>
        public bool CustomStyle { get; set; }

        /// <summary>
        /// Indicate if the style must appear in the Style gallery 
        /// </summary>
        public bool PrimaryStyle { get; set; }

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
        /// Shading color in hex value (RRGGBB format)
        /// </summary>
        public string Shading { get; set; }

        /// <summary>
        /// Font size
        /// </summary>
        public int? FontSize { get; set; }

        /// <summary>
        /// Font color in hex value (RRGGBB format)
        /// </summary>
        public string FontColor { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public Style()
            : base(typeof(Style).Name)
        {
            CustomStyle = true;
        }
    }
}
