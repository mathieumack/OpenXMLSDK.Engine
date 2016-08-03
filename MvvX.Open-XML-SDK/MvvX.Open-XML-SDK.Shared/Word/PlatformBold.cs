using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word;

namespace MvvX.Open_XML_SDK.Shared.Word
{
    public class PlatformBold : PlatformOnOffType, IBold
    {
        private readonly Bold bold;

        public PlatformBold(Bold bold)
            : base(bold)
        {
            this.bold = bold;
        }

        #region Static helpers methods

        public static PlatformBold New()
        {
            return new PlatformBold(new Bold());
        }

        #endregion
    }
}
