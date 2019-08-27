using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenXMLSDK.Engine.Word.ReportEngine.BatchModels;
using OpenXMLSDK.Engine.Word.ReportEngine.Models;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    /// <summary>
    /// Extension class for Table Cell rendering
    /// </summary>
    public static class TemplateModelExtensions
    {
        /// <summary>
        /// Render a cell in a word document
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="parent"></param>
        /// <param name="context"></param>
        /// <param name="documentPart"></param>
        /// <param name="formatProvider"></param>
        /// <returns></returns>
        public static void Render(this TemplateModel templateModel, 
                                        Document document, 
                                        OpenXmlElement parent, 
                                        ContextModel context, 
                                        OpenXmlPart documentPart, 
                                        IFormatProvider formatProvider)
        {
            if (document.TemplateDefinitions == null)
                throw new ArgumentNullException(nameof(document), "There is no template definitions defined in the document");
            if (!document.TemplateDefinitions.Any(e => e.TemplateId == templateModel.TemplateId))
                throw new ArgumentNullException(nameof(document), "the template does not exists in the template definition list");

            var templateDefinition = document.TemplateDefinitions.FirstOrDefault(e => e.TemplateId == templateModel.TemplateId);

            if (templateDefinition.ChildElements == null)
                return ;

            foreach (var child in templateDefinition.ChildElements)
                child.Render(document, parent, context, documentPart, formatProvider);
        }
    }
}
