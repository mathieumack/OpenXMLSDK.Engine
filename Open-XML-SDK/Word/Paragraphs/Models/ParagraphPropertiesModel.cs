using MvvX.Plugins.OpenXMLSDK.Word.Models;

namespace MvvX.Plugins.OpenXMLSDK.Word.Paragraphs.Models
{
    public class ParagraphPropertiesModel
    {
        public NumberingPropertiesModel NumberingProperties { get; set; }

        public ParagraphStyleIdModel ParagraphStyleId { get; set; }

        public SpacingBetweenLinesModel SpacingBetweenLines { get; set; }

        public ParagraphBordersModel ParagraphBorders { get; set; }
    }
}