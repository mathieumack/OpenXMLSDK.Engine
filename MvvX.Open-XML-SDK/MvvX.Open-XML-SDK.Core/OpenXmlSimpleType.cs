namespace MvvX.Open_XML_SDK.Core
{
    public abstract class OpenXmlSimpleType
    {
        public bool HasValue { get; }

        public string InnerText { get; set; }

        protected string TextValue { get; set; }
    }
}
