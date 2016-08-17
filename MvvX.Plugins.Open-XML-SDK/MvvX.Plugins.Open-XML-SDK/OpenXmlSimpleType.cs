namespace MvvX.Plugins.OpenXMLSDK
{
    public abstract class OpenXmlSimpleType
    {
        public bool HasValue { get; }

        public string InnerText { get; set; }

        protected string TextValue { get; set; }
    }
}
