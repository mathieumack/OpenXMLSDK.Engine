using System.Collections.Generic;

namespace OpenXMLSDK.Engine.Word.ReportEngine.Models
{
    /// <summary>
    /// Define a template definition (List of BaseElements)
    /// </summary>
    public class TemplateDefinition
    {
        /// <summary>
        /// Id of the template
        /// </summary>
        public string TemplateId { get; set; }

        /// <summary>
        /// Comment that you can use in the json document. Not used by the Report engine
        /// </summary>
        public string Note { get; set; }

        /// <summary>
        /// List of controls contained in the template
        /// </summary>
        public List<BaseElement> ChildElements { get; set; } = new List<BaseElement>();
    }
}
