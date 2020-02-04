namespace ReportEngine.Core.Template.ExtendedModels
{
    public class UnderlineModel
    {
        /// <summary>
        /// Type of underline
        /// </summary>
        public UnderlineValues Val { get; set; }

        /// <summary>
        /// Color of the under line
        /// Hexa code without #
        /// </summary>
        public string Color { get; set; }
    }
}
