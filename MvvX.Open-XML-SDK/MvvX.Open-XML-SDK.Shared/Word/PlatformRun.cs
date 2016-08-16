using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word;

namespace MvvX.Open_XML_SDK.Shared.Word
{
    public class PlatformRun : PlatformOpenXmlElement, IRun
    {
        private readonly Run run;

        private IRunProperties properties;
        public IRunProperties Properties
        {
            get
            {
                if (properties == null)
                    properties = PlatformRunProperties.New(run);
                return properties;
            }
        }

        public PlatformRun(Run run)
            : base(run)
        {
            this.run = run;
        }

        #region Static helpers methods

        public static PlatformRun New()
        {
            return new PlatformRun(new Run());
        }

        #endregion
    }
}
