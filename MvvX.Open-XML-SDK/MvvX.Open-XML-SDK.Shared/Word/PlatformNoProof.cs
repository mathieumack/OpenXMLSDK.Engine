using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word;

namespace MvvX.Open_XML_SDK.Shared.Word
{
    public class PlatformNoProof : PlatformOnOffType, INoProof
    {
        private readonly NoProof noProof;

        public PlatformNoProof(NoProof noProof)
            : base(noProof)
        {
            this.noProof = noProof;
        }

        #region Static helpers methods

        public static PlatformNoProof New()
        {
            return new PlatformNoProof(new NoProof());
        }

        #endregion
    }
}
