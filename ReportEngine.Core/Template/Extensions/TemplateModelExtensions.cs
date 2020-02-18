using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportEngine.Core.Template.Extensions
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
    }
}
