namespace OpenXMLSDK.Engine.Word.Tables.Models
{
    public class TableLayoutModel
    {
        /// <summary>
        /// Type
        /// default : Autofit
        /// </summary>
        public TableLayoutValues Type { get; set; }

        public TableLayoutModel()
        {
            Type = TableLayoutValues.Autofit;
        }
    }
}
