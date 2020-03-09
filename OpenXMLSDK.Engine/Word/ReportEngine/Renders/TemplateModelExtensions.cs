using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using OpenXMLSDK.Engine.Word.ReportEngine.Models;
using OpenXMLSDK.Engine.Word.ReportEngine.Renders;

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
        public static List<BaseElement> ExtractTemplateItems(this TemplateModel templateModel, 
                                                                    Document document)
        {
            if (document.TemplateDefinitions == null)
                throw new ArgumentNullException(nameof(document), "There is no template definitions defined in the document");
            if (!document.TemplateDefinitions.Any(e => e.TemplateId == templateModel.TemplateId))
                throw new ArgumentNullException(nameof(document), "the template does not exists in the template definition list");

            var templateDefinition = document.TemplateDefinitions.FirstOrDefault(e => e.TemplateId == templateModel.TemplateId);

            if (templateDefinition.ChildElements == null)
                return new List<BaseElement>();

            return templateDefinition.Clone().ChildElements;
        }

        /// <summary>
        /// Replace a template ID by his context definition
        /// </summary>
        /// <param name="templateModel">Template model to update</param>
        /// <param name="contextToUpdate">Context model containing the template ID</param>
        public static void ReplaceItem(this TemplateModel templateModel, ContextModel contextContainingTheID)
        {
            // Search the template ID to use
            var templateIdToFindInCurrentContext = templateModel.TemplateId;
            var templateIdToUse = contextContainingTheID.GetItem<StringModel>(templateIdToFindInCurrentContext)?.Value;
            if (string.IsNullOrWhiteSpace(templateIdToUse))
                return;
                        
            // Update template ID with the ID presents in the context
            templateModel.TemplateId = templateIdToUse;
        }
    }
}
