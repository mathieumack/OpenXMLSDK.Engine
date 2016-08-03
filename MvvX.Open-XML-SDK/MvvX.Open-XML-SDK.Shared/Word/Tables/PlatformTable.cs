using MvvX.Open_XML_SDK.Core.Word.Bases;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MvvX.Open_XML_SDK.Core.Word.Tables
{
    public class PlatformTable : PlatformOpenXmlElement, ITable
    {
        private readonly Table table;

        public PlatformTable(Table table)
            : base(table)
        {
            this.table = table;
        }

        #region Static helpers methods

        public static PlatformTable New()
        {
            return new PlatformTable(new Table());
        }

        #endregion
    }
}
