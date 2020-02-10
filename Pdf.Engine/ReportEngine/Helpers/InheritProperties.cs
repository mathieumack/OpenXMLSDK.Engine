namespace Pdf.Engine.ReportEngine.Helpers
{
    public class InheritProperties
    {
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
        /// Shading configured for paragraph or cell
        /// </summary>
        public string Shading { get; set; }
    }
}
