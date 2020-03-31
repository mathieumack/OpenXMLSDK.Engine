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
        /// Hint text, can be null
        /// <para>Shown when the <see cref="SimpleField"/> is not yet calculated</para>
        /// </summary>
        public Label HintText { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public SimpleField()
            : base(typeof(SimpleField).Name)
        {
        }
    }
}
