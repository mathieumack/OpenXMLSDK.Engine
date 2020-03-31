using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenXMLSDK.Engine.ReportEngine.DataContext;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    public static class SimpleFieldExtensions
    {
        /// <summary>
        /// Render Simplefield
        /// </summary>
        /// <param name="simpleField"></param>
        /// <param name="parent"></param>
        /// <param name="context"></param>
        /// <param name="documentPart"></param>
        /// <param name="formatProvider"></param>
        public static void Render(this Models.SimpleField simpleField,
                                  OpenXmlElement parent,
                                  ContextModel context,
                                  OpenXmlPart documentPart,
                                  IFormatProvider formatProvider)
        {
            context.ReplaceItem(simpleField, formatProvider);

            var field = new DocumentFormat.OpenXml.Wordprocessing.SimpleField()
            {
                Instruction = simpleField.Instruction,
                Dirty = simpleField.IsDirty
            };
            parent.AppendChild(field);

            if (simpleField.HintText != null)
            {
                simpleField.HintText.Render(field, context, documentPart, formatProvider);
            }
        }
    }
}
