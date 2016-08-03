using MvvX.Open_XML_SDK.Core.Word.Bases;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word;

namespace MvvX.Open_XML_SDK.Shared.Word.Bases
{
    public class PlatformRunProperties : PlatformOpenXmlElement, IRunProperties
    {
        private readonly RunProperties run;

        public PlatformRunProperties(RunProperties run)
            : base(run)
        {
            this.run = run;
        }

        #region Interface :

        private IBold bold;
        public IBold Bold
        {
            get
            {
                if (run.Bold == null)
                    run.Bold = new Bold();

                if (bold == null)
                    bold = new PlatformBold(run.Bold);

                return bold;
            }
        }

        private IItalic italic;
        public IItalic Italic
        {
            get
            {
                if (run.Italic == null)
                    run.Italic = new Italic();

                if (italic == null)
                    italic = new PlatformItalic(run.Italic);
                return italic;
            }
        }

        private ICaps caps;
        public ICaps Caps
        {
            get
            {
                if (run.Caps == null)
                    run.Caps = new Caps();

                if (caps == null)
                    caps = new PlatformCaps(run.Caps);
                return caps;
            }
        }

        private IDoubleStrike doubleStrike;
        public IDoubleStrike DoubleStrike
        {
            get
            {
                if (run.DoubleStrike == null)
                    run.DoubleStrike = new DoubleStrike();

                if (doubleStrike == null)
                    doubleStrike = new PlatformDoubleStrike(run.DoubleStrike);
                return doubleStrike;
            }
        }

        private IEmboss emboss;
        public IEmboss Emboss
        {
            get
            {
                if (run.Emboss == null)
                    run.Emboss = new Emboss();

                if (emboss == null)
                    emboss = new PlatformEmboss(run.Emboss);
                return emboss;
            }
        }

        private INoProof noProof;
        public INoProof NoProof
        {
            get
            {
                if (run.NoProof == null)
                    run.NoProof = new NoProof();

                if (noProof == null)
                    noProof = new PlatformNoProof(run.NoProof);
                return noProof;
            }
        }

        private IOutline outline;
        public IOutline Outline
        {
            get
            {
                if (run.Outline == null)
                    run.Outline = new Outline();

                if (outline == null)
                    outline = new PlatformOutline(run.Outline);
                return outline;
            }
        }

        private IShadow shadow;
        public IShadow Shadow
        {
            get
            {
                if (run.Shadow == null)
                    run.Shadow = new Shadow();

                if (shadow == null)
                    shadow = new PlatformShadow(run.Shadow);
                return shadow;
            }
        }

        #endregion

        #region Static helpers methods

        public static PlatformRunProperties New()
        {
            return new PlatformRunProperties(new RunProperties());
        }

        #endregion
    }
}
