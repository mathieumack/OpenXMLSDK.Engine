using System;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Open_XML_SDK.Core.Word;

namespace MvvX.Open_XML_SDK.Shared.Word
{
    public class PlatformGridSpan : PlatformDecimalNumberType, IGridSpan
    {
        private readonly GridSpan gridSpan;

        public PlatformGridSpan(GridSpan gridSpan)
            : base(gridSpan)
        {
            this.gridSpan = gridSpan;
        }

        #region Static helpers methods

        public static PlatformGridSpan New()
        {
            return new PlatformGridSpan(new GridSpan());
        }

        public static PlatformGridSpan New(GridSpan gridSpan)
        {
            if (gridSpan == null)
                gridSpan = new GridSpan();

            return new PlatformGridSpan(gridSpan);
        }

        #endregion
    }
}
