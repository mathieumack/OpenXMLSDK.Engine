using DocumentFormat.OpenXml;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word
{
    public abstract class PlatformOnOffType : PlatformOpenXmlElement
    {
        protected PlatformOnOffType(OpenXmlElement openXmlElement)
            : base(openXmlElement)
        {
        }

        public OnOffValue Val { get; set; }
    }
}
