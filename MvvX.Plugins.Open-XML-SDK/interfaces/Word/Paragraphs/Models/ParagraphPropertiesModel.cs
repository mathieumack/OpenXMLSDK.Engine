using OpenXMLSDK.Engine.Word.Models;

namespace OpenXMLSDK.Engine.Word.Paragraphs.Models
{
    public class ParagraphPropertiesModel
    {
        public NumberingPropertiesModel NumberingProperties { get; set; }

        public ParagraphStyleIdModel ParagraphStyleId { get; set; }

        public SpacingBetweenLinesModel SpacingBetweenLines { get; set; }

        public ParagraphBordersModel ParagraphBorders { get; set; }
    }
}