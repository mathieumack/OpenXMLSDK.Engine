namespace MvvX.Plugins.Open_XML_SDK.Core.Word
{
    public interface IShading : IOpenXmlElement
    {
        string Color { get; set; }

        string Fill { get; set; }

        ThemeColorValues? ThemeColor { get; set; }

        ThemeColorValues? ThemeFill { get; set; }

        string ThemeFillShade { get; set; }

        string ThemeFillTint { get; set; }

        string ThemeShade { get; set; }

        string ThemeTint { get; set; }

        ShadingPatternValues? Val { get; set; }
    }
}
