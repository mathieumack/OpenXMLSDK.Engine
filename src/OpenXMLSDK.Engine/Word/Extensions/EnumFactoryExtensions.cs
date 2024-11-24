using DocumentFormat.OpenXml;

namespace OpenXMLSDK.Engine.Word.Extensions
{
    internal static class EnumFactoryExtensions
    {
        internal static EnumValue<T> ToEnumValue<T>(this T enumeration) where T : struct, IEnumValue, IEnumValueFactory<T>
        {
            return new EnumValue<T>(enumeration);
        }
    }
}
