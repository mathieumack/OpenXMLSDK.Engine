using System;
using System.Collections.Generic;
using ReportEngine.Core.Template.ExtendedModels;

namespace ReportEngine.Core.Template.Text
{
    /// <summary>
    /// Model for a label
    /// </summary>
    public class Label : BaseElement
    {
        /// <summary>
        /// Define if the content is an html content
        /// </summary>
        [Obsolete("Please use HtmlContent object")]
        public bool IsHtml { get; set; }

        /// <summary>
        /// Label content (can contains #key# from context)
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        /// if true, the label contains the page number
        /// </summary>
        public bool IsPageNumber { get; set; }

        /// <summary>
        /// Indicate if the engine must preserve empty space or space before or after text in generation
        /// </summary>
        public SpaceProcessingModeValues SpaceProcessingModeValue { get; set; }

        /// <summary>
        /// Is text contained in current element bold
        /// </summary>
        public bool? Bold { get; set; }

        /// <summary>
        /// Key from context used to determine element boldness
        /// </summary>
        public string BoldKey { get; set; }

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
        public float? FontSize { get; set; }

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

        /// <summary>
        /// Definition of the underline
        /// Can be null if not defined
        /// </summary>
        public UnderlineModel Underline { get; set; }

        /// <summary>
        /// Define transform operation on text before rendering
        /// You can combine multiple operation
        /// Works only for non Html content
        /// </summary>
        public List<LabelTransformOperation> TransformOperations { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public Label()
            : base(typeof(Label).Name)
        {
        }
    }
}
