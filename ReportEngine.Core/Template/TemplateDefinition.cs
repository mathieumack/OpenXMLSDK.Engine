namespace ReportEngine.Core.Template
{
    /// <summary>
    /// Define a template definition (List of BaseElements)
    /// </summary>
    public class TemplateDefinition : BaseElement
    {
        /// <summary>
        /// Id of the template
        /// </summary>
        public string TemplateId { get; set; }

        /// <summary>
        /// Comment that you can use in the json document. Not used by the Report engine
        /// </summary>
        public string Note { get; set; }

        public TemplateDefinition()
            : base(typeof(TemplateDefinition).Name)
        {

        }
    }
}
