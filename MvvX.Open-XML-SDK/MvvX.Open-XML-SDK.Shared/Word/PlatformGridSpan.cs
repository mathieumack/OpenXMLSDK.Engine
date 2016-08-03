using System;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word;
using MvvX.Open_XML_SDK.Core.Word.Bases;

namespace MvvX.Open_XML_SDK.Shared.Word
{
    public class PlatformGridSpan : PlatformDecimalNumberType, IGridSpan
    {
        private readonly GridSpan gridSpan;

        public static PlatformGridSpan New()
        {
            return new PlatformGridSpan(new GridSpan());
        }

        public PlatformGridSpan(GridSpan gridSpan)
            : base(gridSpan)
        {
            this.gridSpan = gridSpan;
        }
    }
}
