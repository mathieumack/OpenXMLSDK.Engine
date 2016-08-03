using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word;

namespace MvvX.Open_XML_SDK.Shared.Word
{
    public class PlatformItalic : PlatformOnOffType, IItalic
    {
        private readonly Italic italic;

        public PlatformItalic(Italic italic)
            : base(italic)
        {
            this.italic = italic;
        }

        #region Static helpers methods

        public static PlatformItalic New()
        {
            return new PlatformItalic(new Italic());
        }

        #endregion
    }
}
