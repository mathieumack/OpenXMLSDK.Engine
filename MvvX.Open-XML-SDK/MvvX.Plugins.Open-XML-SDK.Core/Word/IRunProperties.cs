namespace MvvX.Plugins.Open_XML_SDK.Core.Word
{
    public interface IRunProperties : IOpenXmlElement
    {
        IRunFonts RunFonts { get; }

        string FontSize { get; set; }

        string Color { get; set; }

        bool? Bold { get; set; }

        bool? Italic { get; set; }

        bool? ItalicComplexScript { get; set; }

        bool? Caps { get; set; }

        bool? DoubleStrike { get; set; }

        bool? Emboss { get; set; }

        bool? NoProof { get; set; }

        bool? Outline { get; set; }

        bool? Shadow { get; set; }
    }
}
