namespace ReportEngine.Core.Template.ExtendedModels
{
    public class ParagraphIndentationModel
    {
        /// <summary>
        /// Left Indentation.
        /// </summary>
        public float? Left { get; set; }

        /// <summary>
        /// Left Indentation in Character Units.
        /// </summary>
        public int? LeftChars { get; set; }

        /// <summary>
        /// Right Indentation.
        /// </summary>
        public float? Right { get; set; }

        /// <summary>
        /// Right Indentation in Character Units.
        /// </summary>
        public int? RightChars { get; set; }
    }
}
