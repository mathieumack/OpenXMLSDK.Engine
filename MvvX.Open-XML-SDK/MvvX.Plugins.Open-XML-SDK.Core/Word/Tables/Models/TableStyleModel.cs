namespace MvvX.Plugins.Open_XML_SDK.Core.Word.Tables.Models
{
    public class TableStyleModel
    {
        /// <summary>
        /// Value
        /// default : TableGrid
        /// </summary>
        public string Val { get; set; }

        public TableStyleModel()
        {
            Val = "TableGrid";
        }
    }
}
