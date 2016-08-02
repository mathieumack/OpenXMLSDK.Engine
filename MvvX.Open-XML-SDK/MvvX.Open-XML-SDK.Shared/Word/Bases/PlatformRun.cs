using DocumentFormat.OpenXml.Wordprocessing;

namespace MvvX.Open_XML_SDK.Core.Word.Bases
{
    public class PlatformRun : PlatformOpenXmlElement, IRun
    {
        private readonly Run run;
        private readonly PlatformOpenXmlElement[] childElements;

        public PlatformRun(Run run)
            : base(run)
        {
            this.run = run;
        }

        public PlatformRun()
        {

        }
    }
}
