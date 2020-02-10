using Newtonsoft.Json;
using System;
using ReportEngine.Core.DataContext;
using ReportEngine.Core.Template;
using ReportEngine.Core.Template.Extensions;
using ReportEngine.Core.Template.Text;
using System.Linq;
using it = iTextSharp.text;
using itp = iTextSharp.text.pdf;
using System.Collections.Generic;

namespace Pdf.Engine.ReportEngine.Renders
{
    internal static class BaseElementExtensions
    {
        internal static T Clone<T>(this T element) where T : BaseElement
        {
            return JsonConvert.DeserializeObject<T>(JsonConvert.SerializeObject(element), new JsonSerializerSettings() { Converters = { new JsonContextConverter() } });
        }

        internal static object Render(this BaseElement element,
                                            Document document,
                                            itp.PdfWriter writer,
                                            ContextModel context,
                                            EngineContext ctx,
                                            IFormatProvider formatProvider)
        {
            ctx.Parents.Add(element);
            context.ReplaceItem(element, formatProvider);

            var result = (object)null;
            if (element.Show)
            {
                result = element.RenderItem(document, writer, context, ctx, formatProvider);
            }
            
            if (ctx.Parents.Any())
                ctx.Parents.RemoveAt(ctx.Parents.Count - 1);

            return result;
        }

        private static object RenderItem(this BaseElement element, 
                                            Document document,
                                            itp.PdfWriter writer, 
                                            ContextModel context,
                                            EngineContext ctx,
                                            IFormatProvider formatProvider)
        {
            // Keep this statement order, because of the UniformGrid inherits from Table
            //if (element is ForEach)
            //{
            //    (element as ForEach).Render(document, parent, context, documentPart, formatProvider);
            //}
            if (element is Label)
            {
                return (element as Label).Render(document, writer, context, ctx, formatProvider);
            }
            //else if (element is BookmarkStart)
            //{
            //    (element as BookmarkStart).Render(parent, context, formatProvider);
            //}
            //else if (element is BookmarkEnd)
            //{
            //    (element as BookmarkEnd).Render(parent, context, formatProvider);
            //}
            //else if (element is Hyperlink)
            //{
            //    (element as Hyperlink).Render(parent, context, documentPart, formatProvider);
            //}
            if (element is Paragraph)
            {
                var modelParagraph = (element as Paragraph);
                ctx.UpdateInherits(modelParagraph);
                var paragraph = modelParagraph.Render(document, writer, context, ctx, formatProvider);
                element.AddToParentContainer(ctx, paragraph);
            }
            //else if (element is Image)
            //{
            //    (element as Image).Render(parent, context, documentPart);
            //}
            //else if (element is UniformGrid)
            //{
            //    (element as UniformGrid).Render(document, parent, context, documentPart, formatProvider);
            //}
            //else if (element is Table)
            //{
            //    (element as Table).Render(document, parent, context, documentPart, formatProvider);
            //}
            //else if (element is TableOfContents)
            //{
            //    (element as TableOfContents).Render(documentPart, context);
            //}
            //else if (element is BarModel)
            //{
            //    (element as BarModel).Render(parent, context, documentPart, formatProvider);
            //}
            //else if (element is PieModel)
            //{
            //    (element as PieModel).Render(parent, context, documentPart, formatProvider);
            //}
            //else if (element is HtmlContent)
            //{
            //    (element as HtmlContent).Render(parent, context, documentPart, formatProvider);
            //}

            if (element.ChildElements != null && element.ChildElements.Count > 0)
            {
                for (int i = 0; i < element.ChildElements.Count; i++)
                {
                    var e = element.ChildElements[i];

                    //if (e is TemplateModel)
                    //{
                    //    var elements = (e as TemplateModel).ExtractTemplateItems(document);
                    //    if (i == element.ChildElements.Count - 1)
                    //        element.ChildElements.AddRange(elements);
                    //    else
                    //        element.ChildElements.InsertRange(i + 1, elements);
                    //}
                    //else
                    {
                        e.Render(document, writer, context, ctx, formatProvider);
                    }
                }
            }

            ctx.EndInherits();

            return null;
        }

        /// <summary>
        /// Associe l'élément à son 1er élément IText parent => à utiliser pour un positionnement relatif
        /// </summary>
        internal static void AddToParentContainer(this BaseElement element, EngineContext ctx, it.IElement IElement)
        {
            var nextParent = element; // L'élément lui même peut être un container
            int i = ctx.Parents.Count - 1;
            while (nextParent != null)
            {
                it.IElementListener iElt;
                if (ctx.IElementContainers.TryGetValue(nextParent, out iElt))
                {
                    // Le parent est un IElement (Document...)
                    iElt.Add(IElement);
                    // On réinitialise la pile des parents :
                    return;
                }

                itp.PdfPCell pCell;
                if (ctx.PCellContainers.TryGetValue(nextParent, out pCell))
                {
                    // Le parent est une cellule de tableau
                    pCell.AddElement(IElement);
                    return;
                }
                if (i >= 0)
                    nextParent = ctx.Parents[i];
                else
                    nextParent = null;
                i--;
            }

            // On réinitialise la pile des parents :
            throw new Exception("L'élément racine Document n'a pas été trouvé"); // Ne doit pas arriver
        }
    }
}
