using System.Linq;

namespace OpenXMLSDK.Engine.Word.Paragraphs.Models
{
    public class ParagraphBorderModel
    {
        /// <summary>
        /// Type of borders.
        /// Default : BorderValues.Single
        /// </summary>
        public BorderValues BorderValue { get; set; }

        /// <summary>
        /// Size of the order.
        /// Default : 1
        /// </summary>
        public int? Size { get; set; }

        /// <summary>
        /// Color of the border
        /// Default : Colors.Black
        /// </summary>
        public string Color { get; set; }

        public ParagraphBorderModel()
        {
            Size = null;
            Color = Colors.White;
            BorderValue = BorderValues.None;
        }
    }
}
