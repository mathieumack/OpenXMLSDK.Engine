using DOP = DocumentFormat.OpenXml.Packaging;
using DOW = DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLSDK.Engine.Platform.Word.Extensions;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template.Styles;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    /// <summary>
    /// Extension for style
    /// </summary>
    public static class StyleExtensions
    {
        /// <summary>
        /// Add a style into document
        /// </summary>
        /// <param name="style"></param>
        /// <param name="spart"></param>
        /// <param name="context"></param>
        public static void Render(this Style style, DOP.StyleDefinitionsPart spart, ContextModel context)
        {
            var oxstyle = new DOW.Style()
            {
                Type = style.Type.ToOOxml(),
                CustomStyle = style.CustomStyle,
                StyleId = style.StyleId,
                StyleName = new DOW.StyleName() { Val = style.StyleId },
                PrimaryStyle = new DOW.PrimaryStyle()
                {
                    Val = style.PrimaryStyle ? DOW.OnOffOnlyValues.On : DOW.OnOffOnlyValues.Off
                }
            };

            var srp = new DOW.StyleRunProperties();
            if (style.Bold.HasValue && style.Bold.Value)
                srp.Append(new DOW.Bold());
            if (style.Italic.HasValue && style.Italic.Value)
                srp.Append(new DOW.Italic());
            if (!string.IsNullOrWhiteSpace(style.FontName))
                srp.Append(new DOW.RunFonts() { Ascii = style.FontName, HighAnsi = style.FontName, EastAsia = style.FontName, ComplexScript = style.FontName });
            if (style.FontSize.HasValue)
                srp.Append(new DOW.FontSize() { Val = style.FontSize.Value.ToString() });
            if (!string.IsNullOrWhiteSpace(style.FontColor))
                srp.Append(new DOW.Color() { Val = style.FontColor });
            if (!string.IsNullOrWhiteSpace(style.Shading))
                srp.Append(new DOW.Shading() { Fill = style.Shading });

            if (!string.IsNullOrWhiteSpace(style.StyleBasedOn))
                oxstyle.Append(new DOW.BasedOn() { Val = style.StyleBasedOn });

            oxstyle.Append(srp);

            spart.Styles.Append(oxstyle);
        }
    }
}
