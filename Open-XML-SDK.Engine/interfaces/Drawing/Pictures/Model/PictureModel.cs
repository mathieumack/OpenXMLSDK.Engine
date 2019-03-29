using OpenXMLSDK.Engine.Packaging;

namespace OpenXMLSDK.Engine.Drawing.Pictures.Model
{
    public class PictureModel
    {
        public ImagePartType ImagePartType { get; set; }

        public long? MaxWidth { get; set; }
        
        public long? MaxHeight { get; set; }
    }
}
