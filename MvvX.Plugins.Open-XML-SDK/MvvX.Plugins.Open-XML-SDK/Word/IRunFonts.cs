namespace MvvX.Plugins.OpenXMLSDK.Word
{
    public interface IRunFonts : IOpenXmlElement
    {
        string Ascii { get; set; }

        string ComplexScript { get; set; }

        string EastAsia { get; set; }

        string HighAnsi { get; set; }
    }
}
