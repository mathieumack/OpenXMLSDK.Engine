namespace MvvX.Plugins.Open_XML_SDK.Core
{
    public class EnumValue<T> : OpenXmlSimpleType where T : struct
    {
        public T Value { get; set; }
    }
}
