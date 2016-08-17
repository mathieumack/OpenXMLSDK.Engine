using MvvmCross.Platform;
using MvvmCross.Platform.Plugins;
using MvvX.Plugins.OpenXMLSDK.Platform.Word;
using MvvX.Plugins.OpenXMLSDK.Word;

namespace MvvX.Plugins.OpenXMLSDK.Platform
{
    public class Plugin : IMvxPlugin
    {
        public void Load()
        {
            Mvx.RegisterType<IWordManager>(() =>
            {
                return new WordManager();
            });
        }
    }
}
