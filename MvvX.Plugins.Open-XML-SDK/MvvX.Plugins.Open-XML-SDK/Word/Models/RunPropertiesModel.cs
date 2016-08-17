namespace MvvX.Plugins.OpenXMLSDK.Word.Models
{
    public class RunPropertiesModel
    {
        public RunFontsModel RunFonts { get; set; }

        public bool? Bold { get; set; }

        public bool? Italic { get; set; }

        /// <summary>
        /// Text color
        /// default : black
        /// </summary>
        public string Color { get; set; } = "000000";

        public string FontSize { get; set; }
    }
}
