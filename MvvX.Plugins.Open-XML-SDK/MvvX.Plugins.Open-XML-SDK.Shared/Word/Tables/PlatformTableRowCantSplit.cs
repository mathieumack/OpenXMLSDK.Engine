using DocumentFormat.OpenXml.Wordprocessing;
using MvvX.Plugins.OpenXMLSDK.Word.Tables;
using System;
using System.Linq;

namespace MvvX.Plugins.OpenXMLSDK.Platform.Word.Tables
{
    public class PlatformTableRowCantSplit : PlatformOpenXmlElement, ITableRowCantSplit
    {
        private readonly CantSplit xmlElement;

        public PlatformTableRowCantSplit()
            : this(new CantSplit() { Val = OnOffOnlyValues.Off })
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
                return xmlElement.Val == OnOffOnlyValues.On ? true : false;
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
            CantSplit xmlElement = null;
            if (tableProperties.Descendants<CantSplit>().Any())
                xmlElement = tableProperties.Descendants<CantSplit>().First();
            else
            {
                xmlElement = new CantSplit() { Val = OnOffOnlyValues.Off };
                tableProperties.Append(xmlElement);
            }
            
            return new PlatformTableRowCantSplit(xmlElement);
        }

        #endregion
    }
}
