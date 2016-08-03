using System.Collections;
using System.Collections.Generic;

namespace MvvX.Open_XML_SDK.Core.Word.Bases
{
    public interface IOpenXmlElement : IEnumerable<IOpenXmlElement>, IEnumerable
    {
        object ContentItem { get; }

        T InsertAfterSelf<T>(T newElement) where T : IOpenXmlElement;

        IEnumerable<T> Ancestors<T>() where T : IOpenXmlElement;

        IEnumerable<T> Descendants<T>() where T : IOpenXmlElement;

        void Append<T>(T itemToAppend) where T : IOpenXmlElement;

        void Append(params IOpenXmlElement[] childs);

        void Append(IEnumerable<IOpenXmlElement> childs);
    }
}
