using MvvmCross.Platform;
using MvvmCross.Platform.Plugins;
using MvvX.Plugins.Open_XML_SDK.Core.Word;
using MvvX.Plugins.Open_XML_SDK.Word;

namespace MvvX.Plugins.Open_XML_SDK.Platform
{
    public class Plugin : IMvxPlugin
    {
        public void Load()
        {
            Mvx.RegisterType<IWordManager, WordManager>();
        }
    }
}
