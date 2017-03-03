using MvvX.Plugins.OpenXMLSDK.Packaging;

namespace MvvX.Plugins.OpenXMLSDK.Word.ReportEngine.Models
{
    /// <summary>
    /// Image model
    /// </summary>
    public class Image : BaseElement
    {
        /// <summary>
        /// Type
        /// </summary>
        public ImagePartType ImagePartType { get; set; }

        /// <summary>
        /// Max width
        /// </summary>
        public long? MaxWidth { get; set; }

        /// <summary>
        /// Max height
        /// </summary>
        public long? MaxHeight { get; set; }

        /// <summary>
        /// Path, set null if using content
        /// </summary>
        public string Path { get; set; }

        /// <summary>
        /// Content, set null if using path
        /// </summary>
        public byte[] Content { get; set; }

        /// <summary>
        /// Template key
        /// </summary>
        public string ContextKey { get; set; }
    }
}
