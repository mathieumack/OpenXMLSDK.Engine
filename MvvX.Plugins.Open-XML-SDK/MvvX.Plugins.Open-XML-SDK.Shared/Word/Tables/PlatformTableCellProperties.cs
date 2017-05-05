using System;
using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word;
using MvvX.Plugins.OpenXMLSDK.Word.Tables;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Tables
{
    public class PlatformTableCellProperties : PlatformOpenXmlElement, ITableCellProperties
    {
        private readonly TableCellProperties xmlElement;

        public PlatformTableCellProperties()
            : this(new TableCellProperties())
        {
        }

        public PlatformTableCellProperties(TableCellProperties tableCell)
            : base(tableCell)
        {
            this.xmlElement = tableCell;
        }

        #region Interface :

        private ITableCellBorders tableCellBorders;
        public ITableCellBorders TableCellBorders
        {
            get
            {
                if (tableCellBorders == null)
                    tableCellBorders = PlatformTableCellBorders.New(xmlElement);

                return tableCellBorders;
            }
        }

        private IGridSpan gridSpan;
        public IGridSpan GridSpan
        {
            get
            {
                if (gridSpan == null)
                    gridSpan = PlatformGridSpan.New(xmlElement);
                return gridSpan;
            }
        }

        private ITableCellWidth tableCellWidth;
        public ITableCellWidth TableCellWidth
        {
            get
            {
                if (tableCellWidth == null)
                    tableCellWidth = PlatformTableCellWidth.New(xmlElement);

                return tableCellWidth;
            }
        }

        private IShading shading;
        public IShading Shading
        {
            get
            {
                if (shading == null)
                    shading = PlatformShading.New(xmlElement);

                return shading;
            }
        }

        public OpenXMLSDK.Word.OnOffOnlyValues? NoWrap
        {
            get
            {
                if (xmlElement.NoWrap == null || !xmlElement.NoWrap.Val.HasValue)
                    return null;
                else
                    return (OpenXMLSDK.Word.OnOffOnlyValues)(int)xmlElement.NoWrap.Val.Value;
            }
            set
            {
                if (value == null)
                    xmlElement.NoWrap = null;
                else
                {
                    if (xmlElement.NoWrap == null)
                        xmlElement.NoWrap = new NoWrap();
                    xmlElement.NoWrap.Val = (DocumentFormat.OpenXml.Wordprocessing.OnOffOnlyValues)(int)value;
                }
            }
        }

        private ITableCellMargin tableCellMargin;
        public ITableCellMargin TableCellMargin
        {
            get
            {
                if (tableCellMargin == null)
                    tableCellMargin = PlatformTableCellMargin.New(xmlElement);

                return tableCellMargin;
            }
        }

        #endregion

        #region Static helpers methods

        public static PlatformTableCellProperties New(TableCell tableCell)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<TableCellProperties>(tableCell);
            return new PlatformTableCellProperties(xmlElement);
        }

        #endregion
    }
}
