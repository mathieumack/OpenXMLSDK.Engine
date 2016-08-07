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
                return properties;
            }
        }

        public PlatformRun(Run run)
            : base(run)
        {
            this.run = run;
            this.properties = PlatformRunProperties.New();
            run.Append(Properties.ContentItem as RunProperties);
        }

        #region Static helpers methods

        public static PlatformRun New()
        {
            return new PlatformRun(new Run());
        }

        #endregion
    }
}
