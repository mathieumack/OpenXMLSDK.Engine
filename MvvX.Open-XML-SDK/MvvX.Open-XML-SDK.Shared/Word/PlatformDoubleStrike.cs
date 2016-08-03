using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word;

namespace MvvX.Open_XML_SDK.Shared.Word
{
    public class PlatformDoubleStrike : PlatformOnOffType, IDoubleStrike
    {
        private readonly DoubleStrike doubleStrike;

        public PlatformDoubleStrike(DoubleStrike doubleStrike)
            : base(doubleStrike)
        {
            this.doubleStrike = doubleStrike;
        }

        #region Static helpers methods

        public static PlatformDoubleStrike New()
        {
            return new PlatformDoubleStrike(new DoubleStrike());
        }

        #endregion
    }
}
