using System;
using System.Collections;
using System.Collections.Generic;
using MvvX.Open_XML_SDK.Core.Word.Paragraphs;

namespace MvvX.Open_XML_SDK.Core.Word.Bases
{
    public interface IOpenXmlElement : IEnumerable<IOpenXmlElement>, IEnumerable
    {
        object ContentItem { get; }

        T InsertAfterSelf<T>(T newElement) where T : IOpenXmlElement;

        IEnumerable<T> Ancestors<T>() where T : IOpenXmlElement;
    }
}
