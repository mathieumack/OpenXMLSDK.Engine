using System;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word
{
    public class PlatformGridSpan : PlatformDecimalNumberType, IGridSpan
    {
        private readonly GridSpan gridSpan;

        public PlatformGridSpan()
            : this(new GridSpan())
        {
        }

        public PlatformGridSpan(GridSpan gridSpan)
            : base(gridSpan)
        {
            this.gridSpan = gridSpan;
        }

        #region Static helpers methods

        public static PlatformGridSpan New(TableCellProperties parent)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<GridSpan>(parent);
            return new PlatformGridSpan(xmlElement);
        }

        #endregion
    }
}
