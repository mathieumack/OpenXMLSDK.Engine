namespace ReportEngine.Core.Template.Styles
{
    public class ShadingModel
    {
        public string Color { get; set; }

        public string Fill { get; set; }

        public ThemeColorValues? ThemeColor { get; set; }

        public ThemeColorValues? ThemeFill { get; set; }

        public string ThemeFillShade { get; set; }

        public string ThemeFillTint { get; set; }

        public string ThemeShade { get; set; }

        public string ThemeTint { get; set; }

        public ShadingPatternValues? Val { get; set; }
    }
}
