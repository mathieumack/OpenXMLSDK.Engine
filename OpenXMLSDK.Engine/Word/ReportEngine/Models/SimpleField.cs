namespace OpenXMLSDK.Engine.Word.ReportEngine.Models
{
    /// <summary>
    /// SimpleField
    /// </summary>
    public class SimpleField : BaseElement
    {
        /// <summary>
        /// Instruction
        /// </summary>
        public string Instruction { get; set; }

        /// <summary>
        /// Is Dirty
        /// </summary>
        public bool IsDirty { get; set; }

        /// <summary>
        /// Content text
        /// </summary>
        public Label Text { get; set; } = new Label();

        /// <summary>
        /// Constructor
        /// </summary>
        public SimpleField()
            : base(typeof(SimpleField).Name)
        {
        }
    }
}
