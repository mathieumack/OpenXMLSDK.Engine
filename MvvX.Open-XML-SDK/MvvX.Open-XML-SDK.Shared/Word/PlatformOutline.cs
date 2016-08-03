using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word;

namespace MvvX.Open_XML_SDK.Shared.Word
{
    public class PlatformOutline : PlatformOnOffType, IOutline
    {
        private readonly Outline outline;

        public PlatformOutline(Outline outline)
            : base(outline)
        {
            this.outline = outline;
        }

        #region Static helpers methods

        public static PlatformOutline New()
        {
            return new PlatformOutline(new Outline());
        }

        #endregion
    }
}
