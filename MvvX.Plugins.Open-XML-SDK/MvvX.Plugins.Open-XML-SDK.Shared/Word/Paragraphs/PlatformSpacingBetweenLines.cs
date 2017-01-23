using System;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.Paragraphs;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Paragraphs
{
    public class PlatformSpacingBetweenLines : PlatformOpenXmlElement, ISpacingBetweenLines
    {
        private readonly SpacingBetweenLines spacingBetweenLines;

        public PlatformSpacingBetweenLines()
            : this(new SpacingBetweenLines())
        {
        }

        public PlatformSpacingBetweenLines(SpacingBetweenLines spacing)
            : base(spacing)
        {
            this.spacingBetweenLines = spacing;
        }

        public string After
        {
            get
            {
                return spacingBetweenLines.After?.Value;
            }
            set
            {
                if (spacingBetweenLines.After == null)
                    spacingBetweenLines.After = new DocumentFormat.OpenXml.StringValue();

                spacingBetweenLines.After.Value = new DocumentFormat.OpenXml.StringValue(value);
            }
        }

        public string Before
        {
            get
            {
                return spacingBetweenLines.Before?.Value;
            }
            set
            {
                if (spacingBetweenLines.Before == null)
                    spacingBetweenLines.Before = new DocumentFormat.OpenXml.StringValue();

                spacingBetweenLines.Before.Value = new DocumentFormat.OpenXml.StringValue(value);
            }
        }

        #region Static helpers methods

        public static PlatformSpacingBetweenLines New(ParagraphProperties spacing)
        {
            var spacingBetweenLines = CheckDescendantsOrAppendNewOne<SpacingBetweenLines>(spacing);
            return new PlatformSpacingBetweenLines(spacingBetweenLines);
        }
        
        #endregion
    }
}