namespace MvvX.Open_XML_SDK.Core.Word.Bases
{
    public interface IRunProperties : IOpenXmlElement
    {
        IBold Bold { get; }

        IItalic Italic { get; }
        
        ICaps Caps { get; }

        IDoubleStrike DoubleStrike { get; }

        IEmboss Emboss { get; }

        INoProof NoProof { get; }

        IOutline Outline { get; }

        IShadow Shadow { get; }
    }
}
