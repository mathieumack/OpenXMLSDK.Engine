using System.Collections;
using System.Collections.Generic;

namespace MvvX.Open_XML_SDK.Core.Word.Bases
{
    public interface IOpenXmlElement : IEnumerable<IOpenXmlElement>, IEnumerable
    {
        object ContentItem { get; set; }

        object ChildItem { get; set; }

        T InsertAfterSelf<T>(T newElement) where T : IOpenXmlElement;

        IEnumerable<T> Ancestors<T>() where T : IOpenXmlElement;
        
        void Append(params IOpenXmlElement[] newChildren);

        void Append(IEnumerable<IOpenXmlElement> newChildren);

        IOpenXmlElement AppendChild<T>(T newChild) where T : IOpenXmlElement;
    }
}
