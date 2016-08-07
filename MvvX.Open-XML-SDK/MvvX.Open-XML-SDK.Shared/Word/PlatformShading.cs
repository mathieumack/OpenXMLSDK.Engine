using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word;

namespace MvvX.Open_XML_SDK.Shared.Word
{
    public class PlatformShading : PlatformOpenXmlElement, IShading
    {
        private readonly Shading shading;

        public static PlatformShading New()
        {
            return new PlatformShading(new Shading());
        }

        public PlatformShading(Shading shading)
            : base(shading)
        {
            this.shading = shading;
        }
    }
}
