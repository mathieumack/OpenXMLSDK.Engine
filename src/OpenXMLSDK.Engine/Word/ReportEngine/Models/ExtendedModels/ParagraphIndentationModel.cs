namespace OpenXMLSDK.Engine.Word.ReportEngine.Models.ExtendedModels
{
    public class ParagraphIndentationModel
    {
        /// <summary>
        /// Left Indentation.
        /// </summary>
        public string Left { get; set; }

        /// <summary>
        /// Left Indentation in Character Units.
        /// </summary>
        public int? LeftChars { get; set; }

        /// <summary>
        /// Right Indentation.
        /// </summary>
        public string Right { get; set; }

        /// <summary>
        /// Right Indentation in Character Units.
        /// </summary>
        public int? RightChars { get; set; }
    }
}
