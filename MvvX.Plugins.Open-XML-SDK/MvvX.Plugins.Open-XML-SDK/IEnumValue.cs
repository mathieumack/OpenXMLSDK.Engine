namespace MvvX.Plugins.OpenXMLSDK
{
    public class EnumValue<T> : OpenXmlSimpleType where T : struct
    {
        public T Value { get; set; }
    }
}
