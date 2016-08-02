using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word.Paragraphs;
using MvvX.Open_XML_SDK.Core.Word.Tables;
using MvvX.Open_XML_SDK.Shared.Word.Tables;

namespace MvvX.Open_XML_SDK.Core.Word.Bases
{
    public class PlatformOpenXmlElement : IOpenXmlElement
    {
        private OpenXmlElement openXmlElement;
        private OpenXmlElement[] childElements;

        public PlatformOpenXmlElement(OpenXmlElement openXmlElement)
        {
            this.openXmlElement = openXmlElement;
        }
        public PlatformOpenXmlElement()
        {
        }

        public object ContentItem
        {
            get
            {
                return openXmlElement;
            }
            set
            {
                openXmlElement = value as OpenXmlElement;
            }
        }

        public object ChildItem
        {
            get
            {
                return childElements;
            }
            set
            {
                childElements = value as OpenXmlElement[];
            }
        }

        public IEnumerable<T> Ancestors<T>() where T : IOpenXmlElement
        {
            if (typeof(T) == typeof(IParagraph))
                return openXmlElement.Ancestors<Paragraph>().Select(e => new PlatformParagraph(e)).Cast<T>();
            else if (typeof(T) == typeof(IText))
                return openXmlElement.Ancestors<Text>().Select(e => new PlatformText(e)).Cast<T>();
            else if (typeof(T) == typeof(IRun))
                return openXmlElement.Ancestors<Run>().Select(e => new PlatformRun(e)).Cast<T>();
            else
                return openXmlElement.Ancestors<OpenXmlElement>().Select(e => new PlatformOpenXmlElement(e)).Cast<T>();
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

            return newElement;

        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return openXmlElement.GetEnumerator();
        }

        public void Append(params IOpenXmlElement[] newChildren)
        {
            List<OpenXmlElement> item = new List<OpenXmlElement>();

            foreach (var element in newChildren)
            {
                var result = element.ContentItem as OpenXmlElement;

                if (result != null)
                {
                    item.Add(result);
                }
            }
            openXmlElement.Append(item);
        }

        public void Append(IEnumerable<IOpenXmlElement> newChildren)
        {
            List<OpenXmlElement> item = new List<OpenXmlElement>();

            foreach (var element in newChildren)
            {
                var result = element.ContentItem as OpenXmlElement;

                if (result != null)
                {
                    item.Add(result);
                }
            }
            openXmlElement.Append(item);
        }

        public IOpenXmlElement AppendChild<T>(T newChild) where T : IOpenXmlElement
        {
            var item = newChild.ContentItem as OpenXmlElement;
            if(item == null)
            {
                var items = newChild.ChildItem as OpenXmlElement[];

                OpenXmlElement childResult = null;
                foreach (var child in items)
                {
                    childResult = openXmlElement.AppendChild(child);
                }
                if (childResult != item)
                {
                    return new PlatformOpenXmlElement(childResult);
                }
                else
                    return newChild;
            }
            else
            {
                var result = openXmlElement.AppendChild(item);
                if (result != item)
                {
                    return new PlatformOpenXmlElement(result);
                }
                else
                    return newChild;
            }
        }

    }
}
