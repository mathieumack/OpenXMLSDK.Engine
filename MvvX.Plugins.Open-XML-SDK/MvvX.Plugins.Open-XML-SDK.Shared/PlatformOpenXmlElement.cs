using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.Paragraphs;
using MvvX.Plugins.OpenXMLSDK.Word.Tables;
using MvvX.Plugins.OpenXMLSDK.Word;
using MvvX.Plugins.OpenXMLSDK.Platform.Word.Paragraphs;
using MvvX.Plugins.OpenXMLSDK.Platform.Word.Tables;
using MvvX.Plugins.OpenXMLSDK.Platform.Word;

namespace MvvX.Plugins.OpenXMLSDK.Platform
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

            throw new NotSupportedException("type " + typeof(T).Name + " is not supported yet.");
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

            throw new NotSupportedException("type " + typeof(T).Name + " is not supported yet.");
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

        public T AppendChild<T>(T newChild) where T : IOpenXmlElement
        {
            openXmlElement.AppendChild(newChild.ContentItem as OpenXmlElement);
            // TODO : Check if the result if the newChild item.
            return newChild;
        }

        public T InsertAfter<T>(T newChild, IOpenXmlElement refChild) where T : IOpenXmlElement
        {
            throw new NotImplementedException();
        }

        public T InsertAt<T>(T newChild, int index) where T : IOpenXmlElement
        {
            openXmlElement.InsertAt(newChild.ContentItem as OpenXmlElement, index);
            // TODO : Check if the result is the newChild item.
            return newChild;
        }

        public T InsertBefore<T>(T newChild, IOpenXmlElement refChild)
        {
            throw new NotImplementedException();
        }

        public T PrependChild<T>(T newChild) where T : IOpenXmlElement
        {
            openXmlElement.PrependChild(newChild.ContentItem as OpenXmlElement);
            // TODO : Check if the result is the newChild item.
            return newChild;
        }

        public void RemoveAllChildren()
        {
            openXmlElement.RemoveAllChildren();
        }

        public T RemoveChild<T>(T oldChild) where T : IOpenXmlElement
        {
            openXmlElement.RemoveChild(oldChild.ContentItem as OpenXmlElement);
            // TODO : Check if the result is the oldChild item.
            return oldChild;
        }

        #region protected methods

        protected static T CheckDescendantsOrAppendNewOne<T>(OpenXmlElement parent) where T : OpenXmlElement, new()
        {
            T xmlElement = null;
            if (parent.Descendants<T>().Any())
                xmlElement = parent.Descendants<T>().First();
            else
            {
                xmlElement = new T();
                parent.Append(xmlElement);
            }
            return xmlElement;
        }

        #endregion
    }
}
