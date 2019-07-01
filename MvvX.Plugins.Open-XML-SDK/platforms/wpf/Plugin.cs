using MvvmCross;
using MvvmCross.Plugin;
using OpenXMLSDK.Engine.Word;

namespace OpenXMLSDK.Engine
{
    [MvxPlugin]
    [Preserve(AllMembers = true)]
    public class Plugin : IMvxPlugin
    {
        public void Load()
        {
            Mvx.IoCProvider.RegisterType<IWordManager>(() => new WordManager());
        }
    }
}

