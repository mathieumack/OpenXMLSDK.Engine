using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.Open_XML_SDK.Core.Word;

namespace MvvX.Plugins.Open_XML_SDK.Shared.Word
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

        public PlatformRun()
            : this(new Run())
        {
        }

        public PlatformRun(Run run)
            : base(run)
        {
            this.run = run;
        }
    }
}
