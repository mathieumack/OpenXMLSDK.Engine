using MvvX.Plugins.OpenXMLSDK.Packaging;

namespace MvvX.Plugins.OpenXMLSDK.Drawing.Pictures.Model
{
    public class PictureModel
    {
        public ImagePartType ImagePartType { get; set; }

        public long? MaxWidth { get; set; }
        
        public long? MaxHeight { get; set; }
    }
}
