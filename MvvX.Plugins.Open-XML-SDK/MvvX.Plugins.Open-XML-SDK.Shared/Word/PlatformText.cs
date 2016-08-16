using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.Open_XML_SDK.Core.Word;
using MvvX.Plugins.Open_XML_SDK.Shared.Word.Extensions;

namespace MvvX.Plugins.Open_XML_SDK.Shared.Word
{
    public class PlatformText : PlatformOpenXmlElement, IText
    {
        private Text text;

        public SpaceProcessingModeValues? Space
        {
            get
            {
                return text.Space.ToPlatform();
            }
            set

            {
                text.Space = value.ToOOxml();
            }
        }

        public PlatformText()
            : this(new Text())
        {
        }

        public PlatformText(Text text)
            : base(text)
        {
            this.text = text;
        }

        #region Static helpers methods
        
        public static PlatformText New(string text)
        {
            return new PlatformText(new Text(text));
        }

        public static PlatformText New(string text, SpaceProcessingModeValues preserveSpaces)
        {
            return new PlatformText(new Text(text)
            {
                Space = preserveSpaces.ToOOxml()
            });
        }

        #endregion
    }
}
