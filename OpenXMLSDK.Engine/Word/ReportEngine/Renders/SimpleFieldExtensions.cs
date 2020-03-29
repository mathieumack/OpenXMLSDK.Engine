using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenXMLSDK.Engine.ReportEngine.DataContext;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    public static class SimpleFieldExtensions
    {
        public static DocumentFormat.OpenXml.Wordprocessing.SimpleField Render(this Models.SimpleField simpleField,
                                        OpenXmlElement parent,
                                        ContextModel context,
                                        OpenXmlPart documentPart,
                                        IFormatProvider formatProvider)
        {
            context.ReplaceItem(simpleField, formatProvider);

            DocumentFormat.OpenXml.Wordprocessing.SimpleField field = new DocumentFormat.OpenXml.Wordprocessing.SimpleField() { Instruction = simpleField.Instruction, Dirty = simpleField.IsDirty };
            parent.AppendChild(field);

            simpleField.Text.Render(field, context, documentPart, formatProvider);

            return field;
        }
    }
}
