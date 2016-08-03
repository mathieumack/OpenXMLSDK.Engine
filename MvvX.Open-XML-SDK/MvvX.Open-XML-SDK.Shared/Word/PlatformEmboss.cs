using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word;

namespace MvvX.Open_XML_SDK.Shared.Word
{
    public class PlatformEmboss : PlatformOnOffType, IEmboss
    {
        private readonly Emboss emboss;

        public PlatformEmboss(Emboss emboss)
            : base(emboss)
        {
            this.emboss = emboss;
        }

        #region Static helpers methods

        public static PlatformEmboss New()
        {
            return new PlatformEmboss(new Emboss());
        }

        #endregion
    }
}
