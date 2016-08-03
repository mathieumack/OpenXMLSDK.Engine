using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word.Paragraphs;
using MvvX.Open_XML_SDK.Core.Word.Tables;
using MvvX.Open_XML_SDK.Shared.Word;

namespace MvvX.Open_XML_SDK.Core.Word.Bases
{
    public class PlatformOpenXmlElement : IOpenXmlElement
    {
        private readonly OpenXmlElement openXmlElement;

        public PlatformOpenXmlElement(OpenXmlElement openXmlElement)
        {
            this.openXmlElement = openXmlElement;
        }

        public object ContentItem
        {
            get
            {
                return openXmlElement;
            }
        }

        public IEnumerable<T> Ancestors<T>() where T : IOpenXmlElement
        {
            if(typeof(T) == typeof(IParagraph))
                return openXmlElement.Ancestors<Paragraph>().Select(e => new PlatformParagraph(e)).Cast<T>();
            else if (typeof(T) == typeof(IText))
                return openXmlElement.Ancestors<Text>().Select(e => new PlatformText(e)).Cast<T>();
            else if (typeof(T) == typeof(IRun))
                return openXmlElement.Ancestors<Run>().Select(e => new PlatformRun(e)).Cast<T>();
            else if (typeof(T) == typeof(IGridSpan))
                return openXmlElement.Ancestors<GridSpan>().Select(e => new PlatformGridSpan(e)).Cast<T>();
            else if (typeof(T) == typeof(ITable))
                return openXmlElement.Ancestors<Table>().Select(e => new PlatformTable(e)).Cast<T>();
            else if (typeof(T) == typeof(ITableRow))
                return openXmlElement.Ancestors<TableRow>().Select(e => new PlatformTableRow(e)).Cast<T>();
            else if (typeof(T) == typeof(ITableCell))
                return openXmlElement.Ancestors<TableCell>().Select(e => new PlatformTableCell(e)).Cast<T>();
            else
                return openXmlElement.Ancestors<OpenXmlElement>().Select(e => new PlatformOpenXmlElement(e)).Cast<T>();
        }

        public IEnumerable<T> Descendants<T>() where T : IOpenXmlElement
        {
            if (typeof(T) == typeof(IParagraph))
                return openXmlElement.Descendants<Paragraph>().Select(e => new PlatformParagraph(e)).Cast<T>();
            else if (typeof(T) == typeof(IText))
                return openXmlElement.Descendants<Text>().Select(e => new PlatformText(e)).Cast<T>();
            else if (typeof(T) == typeof(IRun))
                return openXmlElement.Descendants<Run>().Select(e => new PlatformRun(e)).Cast<T>();
            else if (typeof(T) == typeof(IGridSpan))
                return openXmlElement.Descendants<GridSpan>().Select(e => new PlatformGridSpan(e)).Cast<T>();
            else if (typeof(T) == typeof(ITable))
                return openXmlElement.Descendants<Table>().Select(e => new PlatformTable(e)).Cast<T>();
            else if (typeof(T) == typeof(ITableRow))
                return openXmlElement.Descendants<TableRow>().Select(e => new PlatformTableRow(e)).Cast<T>();
            else if (typeof(T) == typeof(ITableCell))
                return openXmlElement.Descendants<TableCell>().Select(e => new PlatformTableCell(e)).Cast<T>();
            else
                return openXmlElement.Descendants<OpenXmlElement>().Select(e => new PlatformOpenXmlElement(e)).Cast<T>();
        }

        public void Append<T>(T objectToAppend) where T : IOpenXmlElement
        {
            openXmlElement.Append(objectToAppend.ContentItem as OpenXmlElement);
        }

        public void Append(params IOpenXmlElement[] childs)
        {
            openXmlElement.Append(childs.Select(e => e.ContentItem as OpenXmlElement));
        }

        public void Append(IEnumerable<IOpenXmlElement> childs)
        {
            openXmlElement.Append(childs.Select(e => e.ContentItem as OpenXmlElement));
        }

        public void Append<T>(IList<T> childs) where T : IOpenXmlElement
        {
            openXmlElement.Append(childs.Select(e => e.ContentItem as OpenXmlElement));
        }

        public IEnumerator<IOpenXmlElement> GetEnumerator()
        {
            return CastEnumerator(openXmlElement.GetEnumerator());
        }

        private IEnumerator<IOpenXmlElement> CastEnumerator(IEnumerator<OpenXmlElement> iterator)
        {
            while (iterator.MoveNext())
            {
                yield return new PlatformOpenXmlElement(iterator.Current);
            }
        }

        public T InsertAfterSelf<T>(T newElement) where T : IOpenXmlElement
        {
            var item = newElement.ContentItem as OpenXmlElement;
            var result = openXmlElement.InsertAfterSelf(item);
            // TODO : Check if result if same as newElement
            // If not :
            //return new PlatformOpenXmlElement(result);
            return newElement;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return openXmlElement.GetEnumerator();
        }

    }
}
