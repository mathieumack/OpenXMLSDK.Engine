using System;
using System.Collections.Generic;
using System.Linq;
using OpenXMLSDK.Engine.ReportEngine.DataContext;
using OpenXMLSDK.Engine.Word.ReportEngine.Models;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Renders
{
    /// <summary>
    /// Extension class for <see cref="TemplateModel"/> rendering
    /// </summary>
    public static class TemplateModelExtensions
    {
        /// <summary>
        /// Render a template in a word document
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="parent"></param>
        /// <param name="context"></param>
        /// <param name="documentPart"></param>
        /// <param name="formatProvider"></param>
        /// <returns></returns>
        public static List<BaseElement> ExtractTemplateItems(this TemplateModel templateModel, Document document)
        {
            if (document.TemplateDefinitions is null)
            {
                throw new ArgumentNullException(nameof(document), "There is no template definitions defined in the document");
            }

            var templateDefinition = document.TemplateDefinitions.FirstOrDefault(e => e.TemplateId == templateModel.TemplateId);
            if (templateDefinition is null)
            {
                throw new ArgumentNullException(nameof(document), "the template does not exist in the template definition list");
            }

            if (templateDefinition.ChildElements == null)
            {
                return new List<BaseElement>();
            }

            return templateDefinition.Clone().ChildElements;
        }

        /// <summary>
        /// Replace a template Id by its context definition
        /// </summary>
        /// <param name="templateModel">Template model to update</param>
        /// <param name="contextToUpdate">Context model containing the template ID</param>
        public static void ReplaceItem(this TemplateModel templateModel, ContextModel contextContainingTheId)
        {
            // Search the template Id to use
            if (contextContainingTheId.TryGetItem(templateModel.TemplateId, out StringModel templateIdItem) && !string.IsNullOrWhiteSpace(templateIdItem.Value))
            {        
                // Update template Id with the Id presents in the context
                templateModel.TemplateId = templateIdItem.Value;
            }
        }
    }
}
