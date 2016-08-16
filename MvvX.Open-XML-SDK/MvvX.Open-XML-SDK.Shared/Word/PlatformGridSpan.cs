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

        public static PlatformGridSpan New(TableCellProperties parent)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<GridSpan>(parent);
            return new PlatformGridSpan(xmlElement);
        }

        #endregion
    }
}
