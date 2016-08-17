using System.Collections;
using System.Collections.Generic;

namespace MvvX.Plugins.OpenXMLSDK.Word
{
    public interface IOpenXmlElement : IEnumerable<IOpenXmlElement>, IEnumerable
    {
        object ContentItem { get; }

        T InsertAfterSelf<T>(T newElement) where T : IOpenXmlElement;

        T AppendChild<T>(T newChild) where T : IOpenXmlElement;

        T InsertAfter<T>(T newChild, IOpenXmlElement refChild) where T : IOpenXmlElement;

        T InsertAt<T>(T newChild, int index) where T : IOpenXmlElement;

        T InsertBefore<T>(T newChild, IOpenXmlElement refChild);

        T PrependChild<T>(T newChild) where T : IOpenXmlElement;

        void RemoveAllChildren();

        T RemoveChild<T>(T oldChild) where T : IOpenXmlElement;

        IEnumerable<T> Ancestors<T>() where T : IOpenXmlElement;

        IEnumerable<T> Descendants<T>() where T : IOpenXmlElement;

        void Append<T>(T itemToAppend) where T : IOpenXmlElement;

        void Append(params IOpenXmlElement[] childs);

        void Append(IEnumerable<IOpenXmlElement> childs);
    }
}
