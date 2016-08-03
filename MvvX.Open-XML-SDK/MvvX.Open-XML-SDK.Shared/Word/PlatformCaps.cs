using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word;

namespace MvvX.Open_XML_SDK.Shared.Word
{
    public class PlatformCaps : PlatformOnOffType, ICaps
    {
        private readonly Caps caps;

        public PlatformCaps(Caps caps)
            : base(caps)
        {
            this.caps = caps;
        }

        #region Static helpers methods

        public static PlatformCaps New()
        {
            return new PlatformCaps(new Caps());
        }

        #endregion
    }
}
