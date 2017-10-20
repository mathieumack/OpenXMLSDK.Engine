using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.Tables;
using System;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Tables
{
    public class PlatformTableRowCantSplit : PlatformOpenXmlElement, ITableRowCantSplit
    {
        private readonly CantSplit xmlElement;

        public PlatformTableRowCantSplit()
            : this(new CantSplit())
        {
        }

        public PlatformTableRowCantSplit(CantSplit cantSplit)
            : base(cantSplit)
        {
            this.xmlElement = cantSplit;
        }

        #region Interface :

        public bool Val
        {
            get
            {
                return (xmlElement.Val == OnOffOnlyValues.On ? true : false);
            }
            set
            {
                xmlElement.Val = value ? OnOffOnlyValues.On : OnOffOnlyValues.Off;
            }
        }

        #endregion

        #region Static helpers methods

        public static PlatformTableRowCantSplit New(TableRowProperties tableProperties)
        {
            var xmlElement = CheckDescendantsOrAppendNewOne<CantSplit>(tableProperties);
            return new PlatformTableRowCantSplit(xmlElement);
        }

        #endregion
    }
}
