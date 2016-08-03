using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word;

namespace MvvX.Open_XML_SDK.Shared.Word
{
    public class PlatformShadow : PlatformOnOffType, IShadow
    {
        private readonly Shadow shadow;

        public PlatformShadow(Shadow shadow)
            : base(shadow)
        {
            this.shadow = shadow;
        }

        #region Static helpers methods

        public static PlatformShadow New()
        {
            return new PlatformShadow(new Shadow());
        }

        #endregion
    }
}
